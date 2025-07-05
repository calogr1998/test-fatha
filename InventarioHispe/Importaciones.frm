VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.5#0"; "DXDBGrid.dll"
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "aBox.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{F7E69521-3C28-11D2-B3E7-00AA00B42B7C}#3.1#0"; "fpTab30.ocx"
Begin VB.Form Importaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Costos Estimados de Importación"
   ClientHeight    =   8745
   ClientLeft      =   3165
   ClientTop       =   1815
   ClientWidth     =   11655
   Icon            =   "Importaciones.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   Begin TabproADOLib.fpTabProADO fpTabProADO1 
      Height          =   4935
      Left            =   60
      TabIndex        =   2
      Top             =   2700
      Width           =   11520
      _Version        =   196609
      _ExtentX        =   20320
      _ExtentY        =   8705
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
      TabsPerRow      =   3
      TabCount        =   3
      OffsetFromClientTop=   -1  'True
      DataFormat      =   ""
      BookCornerGuardWidth=   105
      BookCornerGuardLength=   405
      DataField       =   ""
      DataMember      =   ""
      TabCaption      =   "Importaciones.frx":000C
      PageEarMarkPictureNext=   "Importaciones.frx":0244
      PageEarMarkPicturePrev=   "Importaciones.frx":0260
      EarMarkPictureNext=   "Importaciones.frx":027C
      EarMarkPicturePrev=   "Importaciones.frx":0298
      Begin Threed.SSFrame SSFrame1 
         Height          =   495
         Left            =   -26594
         TabIndex        =   31
         Top             =   -15914
         Width           =   11475
         _Version        =   65536
         _ExtentX        =   20241
         _ExtentY        =   873
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars2 
            Left            =   0
            Top             =   0
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   131082
            ToolBarsCount   =   1
            ToolsCount      =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Tools           =   "Importaciones.frx":02B4
            ToolBars        =   "Importaciones.frx":4225
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   2955
         Left            =   -26354
         TabIndex        =   3
         Top             =   -19394
         Width           =   11085
         _Version        =   65536
         _ExtentX        =   19553
         _ExtentY        =   5212
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Enabled         =   0   'False
         Begin VB.TextBox txtOrdenDespacho 
            Height          =   285
            Left            =   2880
            MaxLength       =   20
            TabIndex        =   8
            Top             =   750
            Width           =   2175
         End
         Begin VB.TextBox txtProforma 
            Height          =   285
            Left            =   2880
            MaxLength       =   20
            TabIndex        =   7
            Top             =   1785
            Width           =   2130
         End
         Begin VB.TextBox txtCertificado 
            Height          =   285
            Left            =   2880
            MaxLength       =   100
            TabIndex        =   6
            Top             =   3495
            Visible         =   0   'False
            Width           =   7935
         End
         Begin VB.ComboBox cmbtipoembarque 
            Height          =   315
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1245
            Width           =   1860
         End
         Begin VB.TextBox txtCompañiaTransporte 
            Height          =   285
            Left            =   8820
            MaxLength       =   20
            TabIndex        =   4
            Top             =   1245
            Width           =   1995
         End
         Begin aBoxCtl.aBox aboFechaProDespacho 
            Height          =   315
            Left            =   8820
            TabIndex        =   9
            Top             =   705
            Width           =   1260
            _ExtentX        =   2223
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
            Text            =   "06/08/2012"
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
            ButtonPicture   =   "Importaciones.frx":42D0
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
         Begin aBoxCtl.aBox aboFecProArrPuerto 
            Height          =   315
            Left            =   2880
            TabIndex        =   10
            Top             =   2325
            Width           =   1260
            _ExtentX        =   2223
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
            Text            =   "06/08/2012"
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
            ButtonPicture   =   "Importaciones.frx":4622
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
         Begin aBoxCtl.aBox aboFechaSalida 
            Height          =   315
            Left            =   8820
            TabIndex        =   11
            Top             =   1740
            Width           =   1260
            _ExtentX        =   2223
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
            Text            =   "06/08/2012"
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
            ButtonPicture   =   "Importaciones.frx":4974
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
         Begin aBoxCtl.aBox aboFecLLegEmbarque 
            Height          =   315
            Left            =   8820
            TabIndex        =   12
            Top             =   2325
            Width           =   1260
            _ExtentX        =   2223
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
            Text            =   "06/08/2012"
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
            ButtonPicture   =   "Importaciones.frx":4CC6
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
         Begin aBoxCtl.aBox aboFecProgInspeccion 
            Height          =   315
            Left            =   2880
            TabIndex        =   13
            Top             =   2910
            Visible         =   0   'False
            Width           =   1260
            _ExtentX        =   2223
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
            Text            =   "06/08/2012"
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
            ButtonPicture   =   "Importaciones.frx":5018
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
         Begin aBoxCtl.aBox aboFecInspeccion 
            Height          =   315
            Left            =   8820
            TabIndex        =   14
            Top             =   2865
            Visible         =   0   'False
            Width           =   1260
            _ExtentX        =   2223
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
            Text            =   "06/08/2012"
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
            ButtonPicture   =   "Importaciones.frx":536A
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
         Begin aBoxCtl.aBox aboEmision 
            Height          =   315
            Left            =   2880
            TabIndex        =   15
            Top             =   180
            Width           =   1260
            _ExtentX        =   2223
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
            Text            =   "06/08/2012"
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
            ButtonPicture   =   "Importaciones.frx":56BC
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
         Begin aBoxCtl.aBox abofechaconfirma 
            Height          =   315
            Left            =   8820
            TabIndex        =   16
            Top             =   180
            Width           =   1260
            _ExtentX        =   2223
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
            Text            =   "06/08/2012"
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
            ButtonPicture   =   "Importaciones.frx":5A0E
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
         Begin VB.Label lblOrdenDespacho 
            Caption         =   "Orden de Despacho"
            Height          =   240
            Left            =   135
            TabIndex        =   29
            Top             =   750
            Width           =   1590
         End
         Begin VB.Label lblProforma 
            Caption         =   "Proforma del Embarcador"
            Height          =   240
            Left            =   135
            TabIndex        =   28
            Top             =   1830
            Width           =   1860
         End
         Begin VB.Label lblCertificado 
            Caption         =   "Certificado de Inspeccion"
            Height          =   240
            Left            =   135
            TabIndex        =   27
            Top             =   3540
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.Label lblFechaDespacho 
            Caption         =   "Fecha Programada de Despacho"
            Height          =   285
            Left            =   6210
            TabIndex        =   26
            Top             =   795
            Width           =   2400
         End
         Begin VB.Label lblFecProgArrPuerto 
            Caption         =   "Fecha Programada de Arribo al Puerto"
            Height          =   240
            Left            =   135
            TabIndex        =   25
            Top             =   2370
            Width           =   2715
         End
         Begin VB.Label lblFechSalida 
            Caption         =   "Fecha de Salida"
            Height          =   195
            Left            =   6210
            TabIndex        =   24
            Top             =   1830
            Width           =   1230
         End
         Begin VB.Label lblFechLlegEmbarque 
            Caption         =   "Fecha de Llegada de Embarque"
            Height          =   240
            Left            =   6210
            TabIndex        =   23
            Top             =   2370
            Width           =   2355
         End
         Begin VB.Label lblFecProgInspeccion 
            Caption         =   "Fecha Programada de Inspeccion"
            Height          =   240
            Left            =   135
            TabIndex        =   22
            Top             =   3000
            Visible         =   0   'False
            Width           =   2445
         End
         Begin VB.Label lblFecInpeccion 
            Caption         =   "Fecha de Inspeccion"
            Height          =   240
            Left            =   6210
            TabIndex        =   21
            Top             =   2955
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.Label lblTipoEmbarque 
            Caption         =   "Tipo de Embarque"
            Height          =   240
            Left            =   135
            TabIndex        =   20
            Top             =   1290
            Width           =   1365
         End
         Begin VB.Label lblCompañia 
            Caption         =   "Compañia de Transporte"
            Height          =   240
            Left            =   6210
            TabIndex        =   19
            Top             =   1290
            Width           =   1815
         End
         Begin VB.Label lblEmision 
            Caption         =   "Fecha de Emision"
            Height          =   240
            Left            =   135
            TabIndex        =   18
            Top             =   210
            Width           =   1365
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Confirmada del Embarcador"
            Height          =   195
            Left            =   6210
            TabIndex        =   17
            Top             =   225
            Width           =   2445
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   4365
         Left            =   -26369
         OleObjectBlob   =   "Importaciones.frx":5D60
         TabIndex        =   30
         Top             =   -19844
         Width           =   11190
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
         Height          =   4275
         Left            =   120
         OleObjectBlob   =   "Importaciones.frx":9526
         TabIndex        =   32
         Top             =   420
         Width           =   11280
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1275
      Left            =   3540
      TabIndex        =   44
      Top             =   -60
      Width           =   8055
      Begin VB.TextBox TxtCodCli 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1200
         MaxLength       =   11
         TabIndex        =   54
         Top             =   840
         Width           =   1380
      End
      Begin VB.TextBox TxtNomCli 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2580
         TabIndex        =   53
         Top             =   840
         Width           =   5325
      End
      Begin VB.TextBox TxtTelPrv 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6570
         TabIndex        =   48
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TxtRucPrv 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1200
         MaxLength       =   11
         TabIndex        =   47
         Top             =   180
         Width           =   1380
      End
      Begin VB.TextBox TxtDirPrv 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1200
         TabIndex        =   46
         Top             =   480
         Width           =   4890
      End
      Begin VB.TextBox TxtNomPrv 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2580
         TabIndex        =   45
         Top             =   180
         Width           =   5325
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   240
         TabIndex        =   55
         Top             =   900
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Telf."
         Height          =   195
         Left            =   6210
         TabIndex        =   51
         Top             =   600
         Width           =   330
      End
      Begin VB.Label Label3 
         Caption         =   "Dirección"
         Height          =   195
         Left            =   240
         TabIndex        =   50
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   240
         TabIndex        =   49
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1275
      Left            =   1860
      TabIndex        =   41
      Top             =   -60
      Width           =   1635
      Begin VB.TextBox txtnumproforma 
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
         Left            =   180
         MaxLength       =   7
         TabIndex        =   42
         Top             =   660
         Width           =   1320
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Nº PROFORMA"
         Height          =   195
         Left            =   240
         TabIndex        =   43
         Top             =   420
         Width           =   1260
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1275
      Left            =   60
      TabIndex        =   37
      Top             =   -60
      Width           =   1755
      Begin VB.TextBox txtnumero 
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
         Left            =   240
         MaxLength       =   7
         TabIndex        =   38
         Top             =   660
         Width           =   1260
      End
      Begin VB.TextBox TXTCODPRO 
         Height          =   285
         Left            =   1680
         TabIndex        =   40
         Top             =   420
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "IMPORTACION"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   300
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   60
      TabIndex        =   33
      Top             =   7620
      Width           =   11535
      Begin Threed.SSPanel PnlFactor 
         DataSource      =   "DataCons"
         Height          =   285
         Left            =   660
         TabIndex        =   34
         Top             =   240
         Width           =   2310
         _Version        =   65536
         _ExtentX        =   4075
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "0.000000000000000000"
         ForeColor       =   -2147483640
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
         Alignment       =   4
      End
      Begin Threed.SSCheck chkcerrar 
         Height          =   240
         Left            =   9660
         TabIndex        =   35
         Top             =   240
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "&Cerrar Importación"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Factor"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   285
         Width           =   450
      End
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tools           =   "Importaciones.frx":F7EF
      ToolBars        =   "Importaciones.frx":17661
   End
   Begin VB.Frame Frame3D1 
      Height          =   1455
      Left            =   60
      TabIndex        =   52
      Top             =   1200
      Width           =   11535
      Begin VB.ComboBox CboImporta 
         Height          =   315
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   240
         Width           =   4035
      End
      Begin VB.ComboBox CboZona 
         Height          =   315
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox TxtRefere 
         Height          =   300
         Left            =   1500
         MaxLength       =   80
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   64
         Top             =   960
         Width           =   9915
      End
      Begin VB.TextBox TxtSerie 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9570
         MaxLength       =   7
         TabIndex        =   62
         Top             =   240
         Width           =   420
      End
      Begin VB.TextBox TxtNumFac 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   10020
         MaxLength       =   15
         TabIndex        =   61
         Top             =   240
         Width           =   1245
      End
      Begin VB.ComboBox CboUm 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   600
         Width           =   915
      End
      Begin VB.TextBox TxtCant 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2460
         TabIndex        =   58
         Text            =   "0.00"
         Top             =   600
         Width           =   915
      End
      Begin aBoxCtl.aBox aboFecha 
         Height          =   285
         Left            =   1500
         TabIndex        =   56
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
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
         Text            =   "06/08/2012"
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
         ButtonPicture   =   "Importaciones.frx":1784D
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
      Begin VB.Label Label14 
         Caption         =   "Importador"
         Height          =   195
         Left            =   3660
         TabIndex        =   68
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Zona"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3660
         TabIndex        =   67
         Top             =   660
         Width           =   450
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         Caption         =   "Referencia"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   180
         TabIndex        =   65
         Top             =   960
         Width           =   780
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Fac./Flete"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8760
         TabIndex        =   63
         Top             =   285
         Width           =   780
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "U.M."
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   180
         TabIndex        =   59
         Top             =   660
         Width           =   450
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Fecha de Cálculo"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   180
         TabIndex        =   57
         Top             =   300
         Width           =   1290
      End
   End
   Begin VB.Label lblImportacion 
      Alignment       =   1  'Right Justify
      Caption         =   "Importación:"
      Height          =   240
      Left            =   180
      TabIndex        =   1
      Top             =   90
      Width           =   870
   End
   Begin VB.Label lblFecha 
      Caption         =   "Fecha:"
      Height          =   240
      Left            =   180
      TabIndex        =   0
      Top             =   540
      Width           =   870
   End
End
Attribute VB_Name = "Importaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TbCabImport1 As ADODB.Recordset
Dim sw_nuevo_doc1 As Boolean
Dim sw_detalle1 As Boolean
Dim sw_nuevo_item1 As Boolean
Dim tempo As ADODB.Connection
Dim sw_cabecera1 As Boolean
Dim TbDetOrden1 As ADODB.Recordset
Dim tbProducto1 As ADODB.Recordset
Dim Precios As Boolean
Dim cvalores As String
Dim cmes As String
Dim falta As String
Dim wcantord    As Integer
Dim amovs_cab(0 To 24)  As a_grabacion
Dim amovs_det(0 To 23)  As a_grabacion
Dim amovs_cab1(0 To 7)  As a_grabacion
Dim amovs_det1(0 To 4)  As a_grabacion
Dim factor As Double
Dim ctipo As String
Dim RSDETALLE As ADODB.Recordset
Dim TbCabOrden1 As ADODB.Recordset
Dim TbDetImport1 As ADODB.Recordset
Dim TbDetTmpImp1 As ADODB.Recordset
Dim rst As ADODB.Recordset
Dim flag As Boolean

Dim Temp As ADODB.Connection
Dim tbcostosimp1 As ADODB.Recordset
Dim TbAgente1 As ADODB.Recordset
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
If KeyAscii = 13 Then
    TxtSerie.SetFocus
End If
End Sub

Private Sub abofechaconfirma_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtOrdenDespacho.SetFocus
End If
End Sub


Private Sub CalculaFlete()
Select Case Trim(CboUm.Text)
Case "KG"
    If IsNumeric(TxtCant.Text) And Val(TxtCant.Text) > 0 Then
        sql = "SELECT TIPO, PESO_DE, PESO_A, FACTOR, FAC_UNID,TAR_PAD, ID_ZONA, COSTO "
        sql = sql & "From IMPORT_TARIFA "
        sql = sql & "WHERE (((TIPO)='" & Left(cmbtipoembarque.Text, 1) & "') "
        sql = sql & "AND ((PESO_DE)<=" & Val(TxtCant.Text) & ") "
        sql = sql & "AND ((PESO_A)>=" & Val(TxtCant.Text) & ") "
        sql = sql & "AND ((ID_ZONA)=" & Val(Left(CboZona.Text, 2)) & "))"
        If Rs.State = 1 Then Rs.Close
        Rs.Open sql, cnn_dbbancos, 3, 1
        If Rs.RecordCount > 0 Then
            If Rs!factor = True Then
                PESO_PADRE = traerCampo("IMPORT_TARIFA", "PESO_DE", "ID_TARIFA", Rs!TAR_PAD, "")
                COSTO_PADRE = traerCampo("IMPORT_TARIFA", "COSTO", "ID_TARIFA", Rs!TAR_PAD, "")
                UNID_EXCEDE = (Val(TxtCant.Text) - Val(PESO_PADRE)) / Val(Rs!FAC_UNID & "")
                VALOR_FLETE = (UNID_EXCEDE * Val(Rs!costo & "")) + COSTO_PADRE
                MsgBox VALOR_FLETE
            Else
                MsgBox Rs!costo
            End If
        Else
            MsgBox "No existe escala registrada para este peso", vbInformation, "Flete"
        End If
    End If
End Select
End Sub
Private Sub CboZona_Click()
CalculaFlete
End Sub

Private Sub dxDBGrid1_OnCheckEditToggleClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Text As String, ByVal State As DXDBGRIDLibCtl.ExCheckBoxState)
If dxDBGrid1.Dataset.State = 2 Then dxDBGrid1.Dataset.Post
End Sub

Private Sub dxDBGrid1_OnCustomDrawCell(ByVal hdc As Long, ByVal Left As Single, ByVal Top As Single, ByVal Right As Single, ByVal Bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
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

Private Sub dxDBGrid2_OnAfterDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction)
    If sw_nuevo_item1 = False Then
        If Action = daInsert Then
            dxDBGrid2.Columns.ColumnByFieldName("F3ITEM1").Value = dxDBGrid2.Dataset.RecordCount + 1
            dxDBGrid2.Columns.ColumnByFieldName("F3CANTIDAD").Value = Format(0)
            dxDBGrid2.Columns.ColumnByFieldName("F3PREFOB").Value = Format(0, "0.0000")
            dxDBGrid2.Columns.ColumnByFieldName("f5advalorem").Value = Format(0, "0.0000")
            dxDBGrid2.Columns.ColumnByFieldName("advalorem").Value = Format(0, "0.0000")
            dxDBGrid2.Columns.ColumnByFieldName("base").Value = Format(0, "0.0000")
            dxDBGrid2.Columns.ColumnByFieldName("F3TOTAL").Value = Format(0, "0.0000")
            dxDBGrid2.Columns.ColumnByFieldName("F3PRECOS").Value = Format(0, "0.0000")
            dxDBGrid2.Columns.ColumnByFieldName("F3MARGEN").Value = Format(0, "0.0000")
            dxDBGrid2.Columns.ColumnByFieldName("F3VALVTA").Value = Format(0, "0.0000")
            dxDBGrid2.Columns.ColumnByFieldName("F3DSCTO").Value = Format(0, "0.0000")
            dxDBGrid2.Columns.ColumnByFieldName("F3VTANET").Value = Format(0, "0.0000")
            dxDBGrid2.Columns.ColumnByFieldName("F3PREUNI").Value = Format(0, "0.0000")
            dxDBGrid2.Columns.ColumnByFieldName("F3FLETE").Value = Format(0, "0.00")
            dxDBGrid2.Columns.ColumnByFieldName("CANTIDAD").Value = Format(0)
            dxDBGrid2.Columns.ColumnByFieldName("F5MANUAL").Value = "S"
            dxDBGrid2.Columns.ColumnByFieldName("F3COSTOTOTAL").Value = ""
            dxDBGrid2.Columns.FocusedIndex = 0
        End If
    End If
End Sub


Private Sub dxDBGrid2_OnBeforeDatasetAction(ByVal Action As DXDBGRIDLibCtl.ExDatasetAction, Allow As Boolean)
    If sw_nuevo_item1 = False Then
        If Action = daInsert Then
            If dxDBGrid2.Dataset.RecordCount > 0 Then
                If Len(Trim(dxDBGrid2.Columns.ColumnByFieldName("F3CODFAB").Value & "")) = 0 And Len(Trim(dxDBGrid2.Columns.ColumnByFieldName("F3CODPRO").Value & "")) = 0 Or ChkCerrar.Value Then
                    Allow = False
                End If
            End If
        End If
        If Action = daDelete Then
            dxDBGrid2.Dataset.Refresh
        End If
    End If
End Sub

Private Sub dxDBGrid2_OnEditButtonClick(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
Select Case dxDBGrid2.Columns.FocusedColumn.FieldName
    Case "f2codprov"
        wtipprov = "E"
        ayuda_proveedores.Show vbModal
        wtipprov = ""
        If Trim(wcodprov) <> "" Then
            dxDBGrid2.Dataset.Edit
            dxDBGrid2.Columns.ColumnByFieldName("f2codprov").Value = wrucprov
            dxDBGrid2.Columns.ColumnByFieldName("f2nomprov").Value = wnomprov
        End If
    Case "F5CODPRO"
'        hlp_productos.Show vbModal
        ayuda_productos.Show vbModal
        If Trim(wcodproducto) <> "" Then
            dxDBGrid2.Dataset.Edit
            dxDBGrid2.Columns.ColumnByFieldName("F5CODPRO").Value = wcodproducto
            dxDBGrid2.Columns.ColumnByFieldName("F3CODFAB").Value = wcodfab
            dxDBGrid2.Columns.ColumnByFieldName("F5NOMPRO").Value = wdesproducto
            dxDBGrid2.Columns.ColumnByFieldName("F5PARTARA").Value = wpartar
            dxDBGrid2.Columns.ColumnByFieldName("F5UNIMED").Value = wmedida
'            dxDBGrid2.Columns.FocusedIndex = 6
        End If
End Select
End Sub

Private Sub dxDBGrid2_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)

 If dxDBGrid2.Columns.FocusedColumn.FieldName = "F5CODPRO" Then
 If Len(Trim(dxDBGrid2.Columns.ColumnByFieldName("F5CODPRO").Value)) > 0 Then
'    If dxDBGrid2.Columns.ColumnByFieldName("F5NOMPRO").Value = "" Then
        wf5codpro = dxDBGrid2.Columns.ColumnByFieldName("F5CODPRO").Value
'    Else
'        wf5codpro = dxDBGrid2.Columns.ColumnByFieldName("F3CODFAB").Value
'    End If
    
    If rst.State = adStateOpen Then rst.Close
    sql = "select f5codfab,f5nompro,f7codmed,f5partara,f5valvta,f5factor,f5codpro,f5marca from if5pla where f5codpro='" & wf5codpro & "'"
    rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
    If Not rst.EOF Then
        dxDBGrid2.Dataset.Edit
        dxDBGrid2.Columns.ColumnByFieldName("F5CODPRO").Value = "" & rst("f5codpro")
        dxDBGrid2.Columns.ColumnByFieldName("F3CODFAB").Value = "" & rst("f5codFAB")
        dxDBGrid2.Columns.ColumnByFieldName("F5NOMPRO").Value = "" & rst("f5nompro")
        dxDBGrid2.Columns.ColumnByFieldName("F5PARTARA").Value = "" & rst("f5partara")
        dxDBGrid2.Columns.ColumnByFieldName("F5UNIMED").Value = "" & rst("f7codmed")
        wmarca = "" & rst("F5MARCA")
        If rsttemp.State = adStateOpen Then rsttemp.Close
        sql = "select f2desmar from ef2marcas where f2codmar='" & wmarca & "'"
        rsttemp.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
        If Not rsttemp.EOF Then
            dxDBGrid2.Columns.ColumnByFieldName("F5CODMARCA").Value = wmarca
            dxDBGrid2.Columns.ColumnByFieldName("F5MARCA").Value = "" & rsttemp("f2desmar")
        End If
        rsttemp.Close
        
        'dxDBGrid2.Columns.ColumnByFieldName("F3PREFOB").Value = "" & rst("f5valvta")
        'dxDBGrid2.Columns.ColumnByFieldName("F5FACTOR").Value = "" & rst("f5factor")
        dxDBGrid2.Dataset.Post
'        dxDBGrid2.Columns.FocusedIndex = 6
    Else
        MsgBox "El Producto no Existe", vbInformation, "Sistema de Logística"
        dxDBGrid2.Dataset.Edit
        dxDBGrid2.Columns.ColumnByFieldName("F5CODPRO").Value = ""
        dxDBGrid2.Columns.ColumnByFieldName("F3CODFAB").Value = ""
        dxDBGrid2.Columns.ColumnByFieldName("F5NOMPRO").Value = ""
        dxDBGrid2.Columns.ColumnByFieldName("F5PARTARA").Value = ""
        dxDBGrid2.Columns.ColumnByFieldName("F5UNIMED").Value = ""
        dxDBGrid2.Columns.ColumnByFieldName("F3PREFOB").Value = "0.0000"
        dxDBGrid2.Columns.ColumnByFieldName("F3TOTAL").Value = "0.0000"
        dxDBGrid2.Columns.ColumnByFieldName("F3CANTIDAD").Value = "0.00"
        dxDBGrid2.Dataset.Post
    End If
    rst.Close
End If
ElseIf dxDBGrid2.Columns.FocusedColumn.FieldName = "f2codprov" Then
    'If wf2codprov <> "" Then
        wf2codprov = dxDBGrid2.Columns.ColumnByFieldName("F2CODPROV").Value
        If Trim(wf2codprov) <> "" Then
            If rst.State = adStateOpen Then rst.Close
            sql = "select f2codprov,f2nomprov,f2newruc from ef2proveedores where f2newruc='" & wf2codprov & "'"
            rst.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
            If Not rst.EOF Then
                dxDBGrid2.Dataset.Edit
                dxDBGrid2.Columns.ColumnByFieldName("f2codprov").Value = "" & rst("f2newruc")
                dxDBGrid2.Columns.ColumnByFieldName("f2nomprov").Value = "" & rst("f2nomprov")
                dxDBGrid2.Dataset.Post
            Else
                MsgBox "El Proovedor no Existe", vbInformation, "Sistema de Logística"
                dxDBGrid2.Dataset.Edit
                dxDBGrid2.Columns.ColumnByFieldName("f2codprov").Value = ""
                dxDBGrid2.Columns.ColumnByFieldName("f2nomprov").Value = ""
                dxDBGrid2.Dataset.Post
            End If
            rst.Close
        End If
    'End If
End If

dxDBGrid2.Dataset.Edit
Calcula_Importaciones dxDBGrid2.Columns.FocusedIndex
dxDBGrid2.Dataset.Post
dxDBGrid2.Dataset.Refresh
Calcula_Costos
End Sub


Private Sub dxDBGrid2_OnEditing(ByVal Node As DXDBGRIDLibCtl.IdxGridNode, Allow As Boolean)
If dxDBGrid2.Columns.FocusedColumn.FieldName = "F5CODPRO" And dxDBGrid2.Columns.ColumnByFieldName("F5MANUAL").Value = "N" Then
    dxDBGrid2.Columns.ColumnByFieldName("F5CODPRO").DisableEditor = True
Else
    dxDBGrid2.Columns.ColumnByFieldName("F5CODPRO").DisableEditor = False
End If
End Sub


Private Sub dxDBGrid2_OnKeyDown(KeyCode As Integer, ByVal Shift As Long)
If KeyCode = 115 Then
    If MsgBox("¿Desea Eliminar el Registro Actual?", vbQuestion + vbYesNo, "Sistema de Logística") = vbYes Then
        sw_nuevo_item = True
        If dxDBGrid2.Dataset.RecNo = 1 Then
            dxDBGrid2.Dataset.Delete
            AdicionaItem
        Else
            dxDBGrid2.Dataset.Delete
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
sql = "SELECT F2CODPROV, F2NEWRUC, F2NOMABREV, F2IMPORTA "
sql = sql & "From EF2PROVEEDORES "
sql = sql & "Where F2IMPORTA = -1 ORDER BY F2NOMABREV"
Rs.Open sql, cnn_dbbancos, 3, 1
CboImporta.Clear
If Rs.RecordCount > 0 Then
    Do While Not Rs.EOF
        CboImporta.AddItem Format(Rs!F2CODPROV, "0000") & Space(3) & Rs!F2NOMABREV
        Rs.MoveNext
    Loop
    CboImporta.ListIndex = 0
End If

End Sub
Private Sub Form_Load()
    Me.Left = 1550
    Me.Top = 700
    
    Set TbCabImport1 = New ADODB.Recordset
    Set rsttemp = New ADODB.Recordset
    CargaUM
    CargaZonas
    CargaImportadores
    
    cmbtipoembarque.AddItem "Avion"
    cmbtipoembarque.AddItem "Barco"
    cmbtipoembarque.ListIndex = 0
    
    sw_ayuda_oc = False
    sw_nuevo_doc1 = True
    sw_detalle1 = False
    sw_cabecera1 = False
    sw_nuevo_item1 = False
    
    
    BASE_TEMPORAL1
    TABLA_TEMPORAL1
    DELETEREC_N "TmpDet_Import", tempo
    dxDBGrid2.Dataset.Refresh
    Conf_Grid1
    dxDBGrid2.Dataset.First
    dxDBGrid2.Columns.FocusedIndex = 0
    
    Setea_Import
    
    
    If Val(wnumimp) > 0 Then
        WNUMERO1 = wnumimp
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
    txtnumero.Text = Format(WNUMERO1, "0000000")
    
    
    sw_nuevo_doc = True
    sw_detalle = False
    sw_nuevo_item = False
    sw_ayuda = False
    
    aboFecha.Value = Format(Date, "dd/mm/yyyy")
    aboEmision.Value = Format(Date, "dd/mm/yyyy")
    abofechaconfirma.Value = Format(Date, "dd/mm/yyyy")
    aboFechaProDespacho.Value = Format(Date, "dd/mm/yyyy")
    aboFecProArrPuerto.Value = Format(Date, "dd/mm/yyyy")
    aboFechaSalida.Value = Format(Date, "dd/mm/yyyy")
    aboFecLLegEmbarque.Value = Format(Date, "dd/mm/yyyy")
    aboFecProgInspeccion.Value = Format(Date, "dd/mm/yyyy")
    aboFecInspeccion.Value = Format(Date, "dd/mm/yyyy")
    
    BASE_TEMPORAL
    TABLA_TEMPORAL
    DELETEREC_N DBTable, Temp
    dxDBGrid1.Dataset.Refresh
    Conf_Grid
    
    dxDBGrid1.Dataset.First
    dxDBGrid1.Columns.FocusedIndex = 0

    'txtnumero.Text = Format(WNUMERO1, "0000000")
    llena_items
    
    Set rst = New ADODB.Recordset
    
    If Val(wnumimp) > 0 Then
        inicio = True
        txtnumero.Text = wnumimp
        txtnumero_KeyPress vbKeyReturn
        inicio = False
    End If
End Sub
Public Sub BASE_TEMPORAL()
Set Temp = New ADODB.Connection

base_temp = "TMP_IMP.MDB"

CON = "Provider=Microsoft.JET.OLEDB.4.0; Data Source=" & wrutatemp & "\" & base_temp & "; Persist Security Info=False"
Temp.Open CON

End Sub

Public Sub TABLA_TEMPORAL()
DBTable = "tmp_costos"
End Sub
Public Sub BASE_TEMPORAL1()
Set tempo = New ADODB.Connection

base_temp = "TMP_IMP.MDB"

CON = "Provider=Microsoft.JET.OLEDB.4.0; Data Source=" & wrutatemp & "\" & base_temp & "; Persist Security Info=False"
tempo.Open CON

End Sub

Public Sub TABLA_TEMPORAL1()
DBTable1 = "TmpDet_Import"
End Sub

Private Sub Conf_Grid1()
    
    With dxDBGrid2.Options
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

dxDBGrid2.Dataset.Active = False

If sw_nuevo_doc1 = False Then
    DELETEREC_N "TmpDet_Import", tempo
    dxDBGrid2.Dataset.Refresh
End If

dxDBGrid2.Dataset.ADODataset.ConnectionString = tempo
dxDBGrid2.Dataset.Active = True
dxDBGrid2.Dataset.Close
dxDBGrid2.Dataset.Open

With dxDBGrid2.Dataset

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

    dxDBGrid2.Dataset.Post
    sw_nuevo_item1 = False

End With
dxDBGrid2.Dataset.Close
dxDBGrid2.Dataset.Open
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


Private Sub dxDBGrid1_OnEdited(ByVal Node As DXDBGRIDLibCtl.IdxGridNode)

    If WTipoCambio = 0 Then
        WTipoCambio = 2.65
    End If
    
    If dxDBGrid1.Columns.FocusedIndex = 4 Then
        'If dxDBGrid1.Columns(4).Value > 0 And dxDBGrid1.Columns(5).Value = 0 Then
        If dxDBGrid1.Columns.ColumnByFieldName("F3SOLES").Value > 0 And dxDBGrid1.Columns.ColumnByFieldName("F3DOLARES").Value = 0 Then
          dxDBGrid1.Dataset.Edit
          'dxDBGrid1.Columns(5).Value = Format(dxDBGrid1.Columns(4).Value / WTipoCambio, "0.00")
          dxDBGrid1.Columns.ColumnByFieldName("F3DOLARES").Value = Format(dxDBGrid1.Columns.ColumnByFieldName("F3SOLES").Value / WTipoCambio, "0.00")
          flag = False
          dxDBGrid1.Dataset.Post
          dxDBGrid1.Columns.FocusedIndex = 6
        End If
    End If
    
    If dxDBGrid1.Columns.FocusedIndex = 5 Then
       If dxDBGrid1.Columns.ColumnByFieldName("F3DOLARES").Value > 0 Then
          dxDBGrid1.Dataset.Edit
          dxDBGrid1.Columns.ColumnByFieldName("F3SOLES").Value = Format(dxDBGrid1.Columns.ColumnByFieldName("F3DOLARES").Value * WTipoCambio, "0.00")
          flag = False
          dxDBGrid1.Dataset.Post
        End If
       
    End If
    'If Importaciones.dxDBGrid2.Columns(12).SummaryFooterValue > 0 Then
    If Importaciones.dxDBGrid2.Columns.ColumnByFieldName("F3TOTAL").SummaryFooterValue > 0 Then
        'factor = dxDBGrid1.Columns.ColumnByFieldName("F3DOLARES").SummaryFooterValue / _
        Val(Format(Importaciones.dxDBGrid2.Columns.ColumnByFieldName("F3TOTAL").SummaryFooterValue, "0.000")) + 1
        
        factor = dxDBGrid1.Columns.ColumnByFieldName("F3DOLARES").SummaryFooterValue / _
                (Val(Format(Importaciones.dxDBGrid2.Columns.ColumnByFieldName("F3TOTAL").SummaryFooterValue, "0.000")) _
                + Val(Format(Importaciones.dxDBGrid2.Columns.ColumnByFieldName("F3FLETE").SummaryFooterValue, "0.000"))) + 1
        
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)

    If sw_ayuda_oc = True Then
        Unload importar_ocompra
    End If

End Sub

Private Sub fpTabProADO1_TabActivate(TabToActivate As Integer)
Dim I As Integer

    If TabToActivate = 0 Then
        If flag = False Then
            GRABACIONES
            flag = True
            
            If Val(Importaciones.dxDBGrid2.Columns.ColumnByFieldName("F3TOTAL").SummaryFooterValue) > 0 Then
                factor = dxDBGrid1.Columns.ColumnByFieldName("F3DOLARES").SummaryFooterValue / _
                (Val(Format(Importaciones.dxDBGrid2.Columns.ColumnByFieldName("F3TOTAL").SummaryFooterValue, "0.000")) _
                + Val(Format(Importaciones.dxDBGrid2.Columns.ColumnByFieldName("F3FLETE").SummaryFooterValue, "0.000"))) + 1
                PnlFactor.Caption = Format(factor, "0.000000000000000")
            Else
                PnlFactor.Caption = "0.000000000000000"
            End If
            
            Calcula_Costos
            If dxDBGrid2.Dataset.RecordCount > 0 Then
                For I = 1 To dxDBGrid2.Dataset.RecordCount
                    dxDBGrid2.Dataset.RecNo = I
                    dxDBGrid2.Dataset.Edit
                    Calcula_Importaciones 7
                    dxDBGrid2.Dataset.Post
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
If ChkCerrar Then
    SSPanel1.Enabled = False
    dxDBGrid1.Enabled = False
Else
    SSPanel1.Enabled = True
    dxDBGrid1.Enabled = True
End If
End Sub

Private Sub Txtcodcli_DblClick()
Call TxtCodCli_KeyUp(113, 0)
End Sub

Private Sub TxtCodCli_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
    ayuda_clientes.Show 1
    Txtcodcli.Text = wruccli
    Txtnomcli.Text = wnomcliprov
End If
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
            sw_nuevo_doc1 = False
        Else
            PROCEDIMIENTO_NUEVO
            sw_nuevo_doc1 = True
        End If
                
'        If Panel3D2.Enabled Then
'            If Not inicio Then
''                txtnumproforma.SetFocus
'            End If
'        End If
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
Dim TBCONSULTA1 As ADODB.Recordset
Set TBCONSULTA1 = New ADODB.Recordset
Dim TBCONSULTA2 As ADODB.Recordset
Set TBCONSULTA2 = New ADODB.Recordset
Set tbcostosimp1 = New ADODB.Recordset
Dim TBCONSULTA21 As ADODB.Recordset
Set TBCONSULTA21 = New ADODB.Recordset

dxDBGrid1.Dataset.Close
DELETEREC_N DBTable, Temp
dxDBGrid1.Dataset.Refresh
dxDBGrid1.Dataset.Open

'SQL = "select * from tb_costeodet where f4numimp='" & txtnumero.Text & "' order by f2codigo"
sql = "SELECT TB_COSTEODET.*, BF9GIN.NOMBRE AS F2DESCRIPCION, IIF(ISNULL(BF9GIN.GRUPOFLUJO),'9999',BF9GIN.GRUPOFLUJO) AS GRUPO, IIF(ISNULL(GRUPOS_FLUJO.NOMBRE),'OTROS GASTOS',GRUPOS_FLUJO.NOMBRE) AS NOMBRE " & _
    "FROM (TB_COSTEODET INNER JOIN BF9GIN ON TB_COSTEODET.F2CODIGO = BF9GIN.CODIGO) LEFT JOIN GRUPOS_FLUJO ON BF9GIN.GRUPOFLUJO = GRUPOS_FLUJO.CODIGO WHERE TB_COSTEODET.f4numimp='" & txtnumero.Text & "'"
If TBCONSULTA1.State = adStateOpen Then TBCONSULTA1.Close
TBCONSULTA1.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not TBCONSULTA1.EOF Then
    ntotsoles = 0#
    Do While Not TBCONSULTA1.EOF
        dxDBGrid1.Dataset.Append
        dxDBGrid1.Dataset.Edit
        dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value = TBCONSULTA1.Fields("F2CODIGO")
'        SQL = "Select * from tb_costosimp where F2CODIGO='" & Trim(TBCONSULTA1.Fields("f2codigo")) & "'"
        
'        If tbcostosimp1.State = adStateOpen Then tbcostosimp1.Close
'        tbcostosimp1.Open SQL, cnn_dbbancos, adOpenDynamic, adLockOptimistic
'        If Not tbcostosimp1.EOF Then
'        END IF
        dxDBGrid1.Columns.ColumnByFieldName("F5NOMPRO").Value = TBCONSULTA1.Fields("F2DESCRIPCION")
        dxDBGrid1.Columns.ColumnByFieldName("GRUPO").Value = TBCONSULTA1.Fields("GRUPO")
        dxDBGrid1.Columns.ColumnByFieldName("NOMBRE").Value = TBCONSULTA1.Fields("NOMBRE")
        dxDBGrid1.Columns.ColumnByFieldName("F3CHECK").Value = True
        dxDBGrid1.Columns.ColumnByFieldName("F3PRESUPUESTO").Value = Format(TBCONSULTA1.Fields("F3PRESUPUESTO"), "###,##0.000")
        dxDBGrid1.Columns.ColumnByFieldName("F3DOLARES").Value = Format(TBCONSULTA1.Fields("F3DOLAR"), "###,##0.000")
        dxDBGrid1.Columns.ColumnByFieldName("F3SOLES").Value = Format(TBCONSULTA1.Fields("F3SOLES"), "###,##0.000")
        
        ntotsoles = ntotsoles + Format(TBCONSULTA1.Fields("F3SOLES"), "###,##0.000")
        
         TBCONSULTA1.MoveNext
    Loop
    dxDBGrid1.Dataset.Edit
    dxDBGrid1.Dataset.Post

    sql = "select * from tb_costeocab where f4numimp='" & txtnumero.Text & "' order by f4numimp"
    If TBCONSULTA2.State = adStateOpen Then TBCONSULTA2.Close
    TBCONSULTA2.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic

    If Not TBCONSULTA2.EOF Then
        WTipoCambio = Format(TBCONSULTA2.Fields("F4TIPCAM"), "#0.000")
    End If

Else
        sql = "SELECT TB_COSTOSIMP.F2CODIGO AS F5CODPRO, BF9GIN.NOMBRE AS F5NOMPRO, IIF(ISNULL(BF9GIN.GRUPOFLUJO),'9999',BF9GIN.GRUPOFLUJO) AS GRUPO, IIF(ISNULL(GRUPOS_FLUJO.NOMBRE),'OTROS GASTOS',GRUPOS_FLUJO.NOMBRE) AS NOMBRE " & _
            "FROM (TB_COSTOSIMP INNER JOIN BF9GIN ON TB_COSTOSIMP.F2CODIGO = BF9GIN.CODIGO) LEFT JOIN GRUPOS_FLUJO ON BF9GIN.GRUPOFLUJO = GRUPOS_FLUJO.CODIGO;"
        If tbcostosimp1.State = adStateOpen Then tbcostosimp1.Close
        tbcostosimp1.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not tbcostosimp1.EOF Then
            Do While Not tbcostosimp1.EOF
                dxDBGrid1.Dataset.Append
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Columns.ColumnByFieldName("F5CODPRO").Value = tbcostosimp1.Fields("F5CODPRO")
                dxDBGrid1.Columns.ColumnByFieldName("F5NOMPRO").Value = tbcostosimp1.Fields("F5NOMPRO")
                dxDBGrid1.Columns.ColumnByFieldName("GRUPO").Value = tbcostosimp1.Fields("GRUPO")
                dxDBGrid1.Columns.ColumnByFieldName("NOMBRE").Value = tbcostosimp1.Fields("NOMBRE")
                dxDBGrid1.Columns.ColumnByFieldName("F3CHECK").Value = True
                dxDBGrid1.Columns.ColumnByFieldName("F3PRESUPUESTO").Value = Format(0, "0.000")
                dxDBGrid1.Columns.ColumnByFieldName("F3SOLES").Value = Format(0, "0.000")
                dxDBGrid1.Columns.ColumnByFieldName("F3DOLARES").Value = Format(0, "0.000")
                tbcostosimp1.MoveNext
            Loop
                dxDBGrid1.Dataset.Edit
                dxDBGrid1.Dataset.Post
        End If
End If
dxDBGrid1.Dataset.First
'dxDBGrid1.Columns.FocusedIndex = 0
'dxDBGrid1.Columns.ColumnByName("GRUPO").GroupIndex = 0
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
   txtproforma.SetFocus
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
Private Sub TxtRucPrv_Change()
    'TxtNomPrv.Text = ""
    If Trim(TxtRucPrv.Text) <> "" And sw_cabecera1 = False Then
        sw_cabecera1 = True
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
        'TxtNomPrv.Text = wnomprov
        Me.MousePointer = 1
        TxtRucPrv_KeyPress 13
    End If

End Sub

Private Sub TxtRucPrv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cargar_datos
End If
End Sub

Public Sub cargar_datos()
Set Tbproveedor1 = New ADODB.Recordset

sql = "Select * from EF2PROVEEDORES where f2newruc='" & Trim(TxtRucPrv.Text) & "'"
If Tbproveedor1.State = adStateOpen Then Tbproveedor1.Close
Tbproveedor1.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
If Not Tbproveedor1.EOF And Len(TxtRucPrv.Text) > 0 Then
    txtcodpro.Text = "" & Tbproveedor1.Fields("F2CODPROV")
    TxtNomPrv.Text = "" & Tbproveedor1.Fields("F2NOMPROV")
    TxtDirPrv.Text = "" & Tbproveedor1.Fields("F2DIRPROV")
    TxtRucPrv.Text = "" & Tbproveedor1.Fields("f2newruc")
    TxtTelPrv.Text = "" & Tbproveedor1.Fields("F2TELPROV")
    aboFecha.SetFocus
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
        TxtNomPrv.Text = "" & Tbproveedor1.Fields("F2NOMPROV")
        TxtDirPrv.Text = "" & Tbproveedor1.Fields("F2DIRPROV")
        txtcodpro.Text = "" & Tbproveedor1.Fields("F2CODPROV")
        TxtRucPrv.Text = "" & Tbproveedor1.Fields("F2NEWRUC")
        TxtTelPrv.Text = "" & Tbproveedor1.Fields("F2TELPROV")
    End If
    Tbproveedor1.Close
    Set Tbproveedor1 = Nothing
    
    txtnumproforma.Text = "" & TbCabImport1.Fields("f4proforma")
    TxtSerie.Text = "" & TbCabImport1.Fields("F4SERIE")
    TxtNumFac.Text = "" & TbCabImport1.Fields("F4NUMFAC")
    aboFecha.Value = "" & Format(TbCabImport1.Fields("F4FECHA"), "DD/MM/YYYY")
    TxtRefere.Text = "" & TbCabImport1.Fields("F4REFERE")
    PnlFactor.Caption = "" & Format(TbCabImport1.Fields("F4FACTOR"), "###,##0.000000000000000000")
    If "" & TbCabImport1.Fields("F4CERRADO") = "S" Then
        ChkCerrar.Value = True
        ChkCerrar.Caption = "Importación Cerrada"
        activar True
    Else
        ChkCerrar.Value = False
        ChkCerrar.Caption = "&Cerrar Importación"
    End If
    
    'Seguimiento
    aboFecInspeccion.Value = "" & TbCabImport1("F4FECINSPE")
    aboFecLLegEmbarque.Value = "" & TbCabImport1("F4FECLLEGADA")
    aboFecProArrPuerto.Value = "" & TbCabImport1("F4FECPUERTO")
    aboEmision.Value = "" & TbCabImport1("F4FECEMISION")
    abofechaconfirma.Value = "" & TbCabImport1("F4FECEMBARCADOR")
    aboFechaProDespacho.Value = "" & TbCabImport1("F4FECDESPACHO")
    aboFechaSalida.Value = "" & TbCabImport1("F4FECSALIDA")
    aboFecProgInspeccion.Value = "" & TbCabImport1("F4FECPROGINSPE")
    
    txtOrdenDespacho.Text = "" & TbCabImport1("F4DESPACHO")
    txtCompañiaTransporte.Text = "" & TbCabImport1("F4TRANSPORTE")
    txtproforma.Text = "" & TbCabImport1("F4PROEMBARCA")
    txtCertificado.Text = "" & TbCabImport1("F4CERTIFICADO")
    
    If "" & TbCabImport1("F4EMBARQUE") <> "" Then
        For I = 0 To cmbtipoembarque.ListCount
            If cmbtipoembarque.List(I) = "" & TbCabImport1("F4EMBARQUE") Then
                cmbtipoembarque.ListIndex = I
            End If
        Next I
    End If
    
    LLena_DataGrid
    sw_cabecera1 = False: sw_detalle1 = False
End Sub

Private Sub LLena_DataGrid()
Dim X As Integer
Dim SQL1 As String
Dim unidad As String

    Set TbDetImport1 = New ADODB.Recordset
    Set tbProducto1 = New ADODB.Recordset

    CONT = 1
    dxDBGrid2.Dataset.Close
    DELETEREC_N "TmpDet_Import", tempo
    SQL1 = "Select F5CODPRO,F5NomPro,f5advalorem,F5FACTOR,f5marca,f5codfab from IF5PLA"
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
        dxDBGrid2.Dataset.ADODataset.ConnectionString = tempo
        dxDBGrid2.Dataset.Active = True
        dxDBGrid2.Dataset.Open
    End If
    tbProducto1.Close
    dxDBGrid2.Dataset.First
    dxDBGrid2.Columns.FocusedIndex = 0
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
    rsif4orden.Open "SELECT * FROM IF4ORDEN WHERE F4NUMORD = '" & (pocompra) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rsif4orden.EOF Then
'--------------------- C. COSTO
'        txtccosto.Text = Trim("" & rsif4orden.Fields("F4CENTRO"))
'        If rsccosto.State = adStateOpen Then rsccosto.Close
'        rsccosto.Open "SELECT F3DESCRIP FROM CENTROS WHERE F3COSTO='" & txtccosto.Text & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
'        If Not rsccosto.EOF Then
'            pnlccosto.Caption = Trim("" & rsccosto.Fields("F3DESCRIP"))
'        End If
'        rsccosto.Close
        '-----------------------------------------------------
'        If Trim("" & rsif4orden.Fields("F4TIPMON")) = "S" Then
'            cmbmoneda.ListIndex = 0
'        ElseIf Trim("" & rsif4orden.Fields("F4TIPMON")) = "D" Then
'            cmbmoneda.ListIndex = 1
'        End If
        Set Tbproveedor1 = New ADODB.Recordset

        sql = "Select * from EF2PROVEEDORES where f2newruc='" & Trim(CODPROV) & "'"
        If Tbproveedor1.State = adStateOpen Then Tbproveedor1.Close
        Tbproveedor1.Open sql, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not Tbproveedor1.EOF And Len(CODPROV) > 0 Then
            txtcodpro.Text = "" & Tbproveedor1.Fields("F2CODPROV")
            TxtNomPrv.Text = "" & Tbproveedor1.Fields("F2NOMPROV")
            TxtDirPrv.Text = "" & Tbproveedor1.Fields("F2DIRPROV")
            TxtRucPrv.Text = "" & Tbproveedor1.Fields("f2newruc")
            TxtTelPrv.Text = "" & Tbproveedor1.Fields("F2TELPROV")
        End If

        If rsif3orden.State = adStateOpen Then rsif3orden.Close
        rsif3orden.Open "SELECT * FROM IF3ORDEN WHERE F4NUMORD = '" & (pocompra) & "' AND F3CANFAL > 0 ", cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsif3orden.EOF Then
            sw_Orden = True
            dxDBGrid2.Dataset.ADODataset.ConnectionString = tempo
            dxDBGrid2.Dataset.Active = True
        
            dxDBGrid2.Dataset.Close
            dxDBGrid2.Dataset.Open
                        
            dxDBGrid2.OptionEnabled = False
            dxDBGrid2.Dataset.DisableControls
            With dxDBGrid2.Dataset
                I = IIf(dxDBGrid2.Count > 1, dxDBGrid2.Count + 1, 1)
                sw_nuevo_item = True
                rsif3orden.MoveFirst
                Do While Not rsif3orden.EOF
                    If sw_detalle = False Then
                        .Edit
                    Else
                        .Append
                    End If
                    .FieldValues("F3ITEM1") = I
                    .FieldValues("F3NumOrd") = Trim("" & rsif3orden.Fields("F4NUMORD"))
                    .FieldValues("F5CODPRO") = Trim("" & rsif3orden.Fields("F3CODPRO"))
                    .FieldValues("F3CODFAB") = Trim("" & rsif3orden.Fields("F3CODFAB"))
                    
                    If rsif5pla.State = adStateOpen Then rsif5pla.Close
                    rsif5pla.Open "SELECT F5NOMPRO,F5MARCA,F5CODFAB,F5VALVTA,F5PARTARA,F5ADVALOREM FROM IF5PLA WHERE F5CODPRO = '" & Trim("" & rsif3orden.Fields("F3CODPRO")) & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
                    If Not rsif5pla.EOF Then
                        .FieldValues("F5NOMPRO") = "" & rsif5pla.Fields("F5NOMPRO")
                        .FieldValues("F3ValVta") = Val("" & rsif5pla.Fields("f5valvta"))
                        .FieldValues("F5CODMARCA") = "" & rsif5pla.Fields("f5marca")
                        .FieldValues("F5PARTARA") = "" & rsif5pla.Fields("f5partara")
                        ADVALOREM = Val("" & rsif5pla.Fields("f5advalorem"))
                    End If
                    rsif5pla.Close
                    .FieldValues("F5MARCA") = "" & rsif3orden.Fields("F5marca")
                    .FieldValues("F5UniMed") = "" & rsif3orden.Fields("UNIDAD")
                    .FieldValues("F3CANTIDAD") = Format(Val("" & rsif3orden.Fields("F3CANFAL")), "###,##0.0000")
                    .FieldValues("F3PREFOB") = Format(Val("" & rsif3orden.Fields("F3Precos")), "###,##0.0000")
                    .FieldValues("F3TOTAL") = Format(Val("" & rsif3orden.Fields("F3Total")), "###,##0.00")
                    .FieldValues("F3PreCos") = Format(0, "###,##0.0000")
                    .FieldValues("F3MARGEN") = Format(0, "###,##0.0000")
                    .FieldValues("F3dscto") = Format(0, "###,##0.0000")
                    .FieldValues("CANTIDAD") = Format(Val("" & rsif3orden.Fields("F3CANFAL")), "###,##0.0000")
                    .FieldValues("F5MANUAL") = "N"
                    .FieldValues("advalorem") = Format(ADVALOREM * Val("" & rsif3orden.Fields("F3Precos")), "###,##0.0000")
                    .FieldValues("base") = Format(Val(Format(rsif3orden.Fields("F3Precos"), "0.0000")) + Format(ADVALOREM * Val(Format(rsif3orden.Fields("F3Precos"), "0.0000")), "0.0000"), "0.0000")
                    .FieldValues("f2codprov") = ""
                    .FieldValues("f2nomprov") = ""
                                                            
                    WTOTAL = WTOTAL + Val(Format(rsif3orden.Fields("F3Total"), "0.0000"))
                    
                    dxDBGrid2.Dataset.Edit
                    sw_detalle = True
                    Calcula_Importaciones 7
                    dxDBGrid2.Dataset.Post
                    dxDBGrid2.Dataset.First
                    dxDBGrid2.Columns.FocusedIndex = 0
                    
                    rsif3orden.MoveNext
'Calcula pvunitario
'                    dxDBGrid1.Dataset.Edit
'                    sw_detalle = True
'                    If dxDBGrid2.Columns.ColumnByFieldName("AFECTO").Value = "*" Then
'                        dxDBGrid2.Columns.ColumnByFieldName("IGV").Value = Format((Val(Format(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value, "0.00")) * Val(Format(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value, "0.0000"))) * (wwigv / 100), "###,###,##0.00")
'                        dxDBGrid2.Columns.ColumnByFieldName("PVUNIT").Value = Format(Val(Format(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value, "0.0000")) * (1 + wwigv / 100), "###,###,##0.0000")
'                    Else
'                        dxDBGrid1.Columns.ColumnByFieldName("IGV").Value = Format(0, "0.00")
'                    End If
'                    dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").Value = dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value * dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value
'                    dxDBGrid1.Columns.ColumnByFieldName("TOTAL").Value = Format(Val(Format(dxDBGrid1.Columns.ColumnByFieldName("CANTIDAD").Value, "0.00")) * Val(Format(dxDBGrid1.Columns.ColumnByFieldName("COSTOUNI").Value, "0.0000")) + Val(Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").Value, "0.00")), "###,###,##0.00")
'                    dxDBGrid2.Dataset.Post
                    I = I + 1
                Loop
                .Edit
                .Post

                sw_nuevo_item = False
            End With
            dxDBGrid2.Dataset.EnableControls
            dxDBGrid2.Dataset.Close
            dxDBGrid2.Dataset.Open
            dxDBGrid2.OptionEnabled = True
'            txtTotvv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("VVTOTAL").SummaryFooterValue, "#,###,###,##0.00")
'            txtTotigv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("IGV").SummaryFooterValue, "#,###,###,##0.00")
'            txtTotpv.Text = Format(dxDBGrid1.Columns.ColumnByFieldName("TOTAL").SummaryFooterValue, "#,###,###,##0.00")
            
            
'           Acumula los codigos de orden de compra en el arreglo wnumsord
'            txtOcompra.Text = Val(pocompra)
'            For j = 0 To 999
'                If wnumsord(j) = 0 Then
'                        wnumsord(j) = Val(pocompra)
'                        j = 999
'                End If
'            Next
            
            SSActiveToolBars1.Tools.ITEM("ID_Grabar").Enabled = True
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
    
    xcont = dxDBGrid2.Dataset.RecordCount
    xfab = dxDBGrid2.Dataset.FieldValues("f3codfab")
    dxDBGrid2.Dataset.Close
    
    If xcont = 1 And xfab = "" Then
        DELETEREC_N "TmpDet_Import", tempo
    End If
    Set rst = New ADODB.Recordset
    With dxDBGrid2.Dataset
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
        dxDBGrid2.Dataset.ADODataset.ConnectionString = tempo
        dxDBGrid2.Dataset.Active = True
        dxDBGrid2.Dataset.Open
    End With
    'rsproduc.Close
    dxDBGrid2.Dataset.Edit
    Calcula_Importaciones 7
    dxDBGrid2.Dataset.Post
    dxDBGrid2.Dataset.First
    dxDBGrid2.Columns.FocusedIndex = 0
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
    TxtRefere.SetFocus
End If
End Sub

Private Sub TxtRefere_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    dxDBGrid2.SetFocus
    dxDBGrid2.Columns.FocusedIndex = 0
End If
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
            .TxtFecha.Text = aboFecha.Value
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
        
        Me.MousePointer = 11
        dxDBGrid2.Dataset.Edit
        If dxDBGrid2.Dataset.State = dsEdit Or dxDBGrid2.Dataset.State = dsInsert Then
             dxDBGrid2.Dataset.Post
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
'        RegImporta2.Caption = "Registro de Importaciones"
'        RegImporta2.DataControl.ConnectionString = tempo
'        RegImporta2.DataControl.Source = SQL
'        RegImporta2.Label1.Caption = wempresa
'        RegImporta2.Label27.Caption = aboFecha.Value
'        RegImporta2.Label28.Caption = Importaciones.TxtNomPrv.Text
'        RegImporta2.Label29.Caption = Importaciones.TxtDirPrv.Text
'        RegImporta2.Label30.Caption = Importaciones.TxtTelPrv.Text
'        RegImporta2.Label31.Caption = Importaciones.TxtRefere.Text
'        RegImporta2.Show vbModal
'        Me.MousePointer = vbDefault
    
    Case "idtraduccion"
        Me.MousePointer = vbHourglass
        sql = "Select * from TmpDet_Import order by f5marca"
        RegImporta.DataControl1.ConnectionString = tempo
        RegImporta.DataControl1.Source = sql
        RegImporta.Caption = "Traducción"
        RegImporta.Label1.Caption = wempresa
        RegImporta.txtproforma.Text = Importaciones.txtnumproforma.Text
        RegImporta.Label27.Caption = aboFecha.Value
        RegImporta.Label28.Caption = Importaciones.TxtNomPrv.Text
        RegImporta.Label29.Caption = Importaciones.TxtDirPrv.Text
        RegImporta.Label30.Caption = Importaciones.TxtTelPrv.Text
        RegImporta.Label31.Caption = Importaciones.TxtRefere.Text
        RegImporta.Show vbModal
        Me.MousePointer = vbDefault
        
    Case "ID_Borrar"
        Me.MousePointer = 11
        elimina txtnumero.Text
        sw_nuevo_doc1 = True
        factor = Format(0, "0.000000000000000000")
        nuevo
        dxDBGrid2.Dataset.Close
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
        If dxDBGrid2.Dataset.State = dsEdit Then
            dxDBGrid2.Dataset.Post
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
            cnn_dbbancos.Execute csql
            'AlmacenaQuery_sql csql, cnn_dbbancos
            csql = ("DELETE * FROM TB_COSTEODET WHERE F4NUMIMP='" & txtnumero.Text & "'")
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
nuevo
AdicionaItem1
AdicionaItem1
llena_items
sw_nuevo_doc1 = True
Me.MousePointer = 1
End Sub

Public Sub nuevo()
TxtRucPrv.Text = ""
TxtNomPrv.Text = ""
TxtDirPrv.Text = ""
TxtTelPrv.Text = ""
aboFecha.Value = Format(Now, "dd/mm/yyyy")
TxtSerie.Text = ""
TxtNumFac.Text = ""
TxtRefere.Text = ""
PnlFactor.Caption = Format(0, "0.000000000000000000")
txtnumproforma.Text = ""
TxtRucPrv.SetFocus
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
    
    '---------------------ASIGNA DATOS A LA CABECERA DE IMPORT_CAB
    amovs_cab(0).campo = "F4NUMIMP": amovs_cab(0).valor = txtnumero.Text: amovs_cab(0).TIPO = "T"
    amovs_cab(1).campo = "F4CODPRV": amovs_cab(1).valor = Trim(txtcodpro.Text): amovs_cab(1).TIPO = "T"
    amovs_cab(2).campo = "F4FECHA": amovs_cab(2).valor = aboFecha.Value: amovs_cab(2).TIPO = "F"
    amovs_cab(3).campo = "F4REFERE": amovs_cab(3).valor = IIf(Len(Trim(TxtRefere.Text)) = 0, " ", TxtRefere.Text): amovs_cab(3).TIPO = "T"
    amovs_cab(4).campo = "F4FACTOR": amovs_cab(4).valor = Val(Format(PnlFactor.Caption, "0.000000000000000000")): amovs_cab(4).TIPO = "N"
    amovs_cab(5).campo = "F4SERIE": amovs_cab(5).valor = IIf(Len(Trim(TxtSerie.Text)) = 0, " ", Left(TxtSerie.Text, 3)): amovs_cab(5).TIPO = "T"
    amovs_cab(6).campo = "F4NUMFAC": amovs_cab(6).valor = IIf(Len(Trim(TxtSerie.Text)) = 0, " ", Mid(TxtNumFac.Text, 1, 7)): amovs_cab(6).TIPO = "T"
    amovs_cab(7).campo = "F4TOTAL": amovs_cab(7).valor = dxDBGrid2.Columns(9).SummaryFooterValue: amovs_cab(7).TIPO = "N"
    amovs_cab(8).campo = "F4CERRADO": amovs_cab(8).valor = IIf(ChkCerrar.Value, "S", "N"): amovs_cab(8).TIPO = "T"
    amovs_cab(9).campo = "F4PROFORMA": amovs_cab(9).valor = txtnumproforma.Text: amovs_cab(9).TIPO = "T"
    
    amovs_cab(10).campo = "F4FECINSPE": amovs_cab(10).valor = "" & aboFecInspeccion.Value: amovs_cab(10).TIPO = "F"
    amovs_cab(11).campo = "F4FECLLEGADA": amovs_cab(11).valor = "" & aboFecLLegEmbarque.Value: amovs_cab(11).TIPO = "F"
    amovs_cab(12).campo = "F4FECPUERTO": amovs_cab(12).valor = "" & aboFecProArrPuerto.Value: amovs_cab(12).TIPO = "F"

    amovs_cab(13).campo = "F4FECEMISION": amovs_cab(13).valor = "" & aboEmision.Value: amovs_cab(13).TIPO = "F"
    amovs_cab(14).campo = "F4FECEMBARCADOR": amovs_cab(14).valor = "" & abofechaconfirma.Value: amovs_cab(14).TIPO = "F"
    amovs_cab(15).campo = "F4FECDESPACHO": amovs_cab(15).valor = "" & aboFechaProDespacho.Value: amovs_cab(15).TIPO = "F"
    amovs_cab(16).campo = "F4FECSALIDA": amovs_cab(16).valor = "" & aboFechaSalida.Value: amovs_cab(16).TIPO = "F"
    amovs_cab(17).campo = "F4FECPROGINSPE": amovs_cab(17).valor = "" & aboFecProgInspeccion.Value: amovs_cab(17).TIPO = "F"

    amovs_cab(18).campo = "F4DESPACHO": amovs_cab(18).valor = "" & txtOrdenDespacho.Text: amovs_cab(18).TIPO = "T"
    amovs_cab(19).campo = "F4TRANSPORTE": amovs_cab(19).valor = "" & txtCompañiaTransporte.Text: amovs_cab(19).TIPO = "T"
    amovs_cab(20).campo = "F4PROEMBARCA": amovs_cab(20).valor = "" & txtproforma.Text: amovs_cab(20).TIPO = "T"
    amovs_cab(21).campo = "F4CERTIFICADO": amovs_cab(21).valor = "" & txtCertificado.Text: amovs_cab(21).TIPO = "T"
    
    amovs_cab(22).campo = "F4EMBARQUE": amovs_cab(22).valor = "" & cmbtipoembarque.Text: amovs_cab(22).TIPO = "T"
    amovs_cab(23).campo = "F4importador": amovs_cab(23).valor = "" & Left(Me.CboImporta.Text, 4): amovs_cab(23).TIPO = "T"
    amovs_cab(24).campo = "F4cliente": amovs_cab(24).valor = "" & traerCampo("ef2clientes", "f2codcli", "f2newruc", Me.Txtcodcli.Text, ""): amovs_cab(24).TIPO = "T"


    sql = "Select * from IF4ORDEN"
    If TbCabOrden1.State = adStateOpen Then TbCabOrden1.Close
    TbCabOrden1.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
    
    sql = "Select * from Import_Det"
    If TbDetImport1.State = adStateOpen Then TbDetImport1.Close
    TbDetImport1.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
    
    sql = "Select * from TmpDet_Import"
    If TbDetTmpImp1.State = adStateOpen Then TbDetTmpImp1.Close
    TbDetTmpImp1.Open sql, tempo, adOpenStatic, adLockOptimistic
    
    intcont = 0
    wnumord = dxDBGrid2.Dataset.FieldValues("F3NUMORD")
    falta = "0"
    I% = 0
    Precios = True
    
    If wf1visualiza_import_venta = "*" Then
        Precios = False
    End If
    
    If TbDetTmpImp1.RecordCount >= 1 Then
        graba = True
        TbDetTmpImp1.MoveFirst
    Else
        graba = False
        TbDetImport1.Close
        TbCabOrden1.Close
        TbDetTmpImp1.Close
        Exit Sub
    End If
    
    Do While Not TbDetTmpImp1.EOF
        graba = True
        If Len(Trim("" & TbDetTmpImp1!F5NOMPRO)) = 0 Then
            graba = False
        End If
        
'        If Len(Trim("" & TbDetTmpImp1!F2NOMPROV)) = 0 Then
'            graba = False
'        End If
    
        If graba Then
            wf3codfab = "" & TbDetTmpImp1!F3CODFAB
            wf5codmarca = "" & TbDetTmpImp1!F5CODMARCA
            'SQL = "Select * from IF3ORDEN WHERE F4NUMORD = " & Val(Format(TbDetTmpImp1!F3NUMORD, "0000000")) & " and F3CODPRO = '" & Trim(TbDetTmpImp1!F5CODPRO) & "'"
            sql = "Select * from IF3ORDEN WHERE F4NUMORD = '" & TbDetTmpImp1!F3NUMORD & "' and " _
            & "f3codfab = '" & wf3codfab & "' and f5codmarca='" & wf5codmarca & "'"
            
            If TbDetOrden1.State = adStateOpen Then TbDetOrden1.Close
            TbDetOrden1.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
            '
            If Not TbDetOrden1.EOF Then
                If ctipo = "A" Then
                    CANTFALTAN = TbDetOrden1.Fields("F3CANFAL") - Val(Format(TbDetTmpImp1!F3CANTIDAD, "0.000"))
                Else
                    CANTFALTAN = TbDetOrden1.Fields("F3CANFAL") + Val(Format(TbDetTmpImp1!cantidad, "0.000")) - Val(Format(TbDetTmpImp1!F3CANTIDAD, "0.0000"))
                End If
                
                'CSQL1 = "UPDATE IF3ORDEN SET F3CANFAL=" & CANTFALTAN & " WHERE F4NUMORD = " & Val(Format(TbDetTmpImp1!F3NUMORD, "0000000")) & " and F3CODPRO = '" & Trim(TbDetTmpImp1!F5CODPRO) & "'"
                CSQL1 = "UPDATE IF3ORDEN SET F3CANFAL=" & CANTFALTAN & " WHERE F4NUMORD = " & _
                Val(Format(TbDetTmpImp1!F3NUMORD, "0000000")) & " and f3codfab = '" & wf3codfab & "' and " _
                & "f5codmarca='" & wf5codmarca & "'"
                
                cnn_dbbancos.Execute (CSQL1)
                'AlmacenaQuery_sql CSQL1, cnn_dbbancos
                
                If wimporta(I%).f4falta = "0" Then
                   If Val(Format(TbDetOrden1.Fields("F3CANFAL"), "0.000")) <> 0 Then
                      falta = "1"
                   End If
                End If
    
                If wnumord <> TbDetTmpImp1!F3NUMORD Then
                    wtotord = 0#
                    sql = "Select * from IF3ORDEN where F4NUMORD=" & Val(Format(wnumord, "0000000")) & ""
                    '
                    If TbDetOrden1.State = adStateOpen Then TbDetOrden1.Close
                    TbDetOrden1.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
                    If Not TbDetOrden1.EOF Then
                        Do While TbDetOrden1.Fields("F4NUMORD") = Val(Format(wnumord, "0000000"))
                            wtotord = wtotord + Val(Format(TbDetOrden1.Fields("F3TOTAL"), "0.0000"))
                            TbDetOrden1.MoveNext
                            If TbDetOrden1.EOF Then Exit Do
                        Loop
    
                        TbCabOrden1.Find "F4NUMORD=" & Val(Format(wnumord, "0000000")) & "", , adSearchForward
                        If Not TbCabOrden1.EOF Then
                            MONTO1 = Val(Format(wtotord, "0.0000"))
                            If falta = "0" And wimporta(I%).f4falta = "0" Then
                                FALTA1 = "0"
                            Else
                                FALTA1 = "1"
                            End If
                            
                            Csql2 = "UPDATE IF4ORDEN SET F4MONTO=" & MONTO1 & ",F4FALTA='" & FALTA1 & "' WHERE F4NUMORD=" & Val(Format(wnumord, "0000000")) & ""
                            cnn_dbbancos.Execute (Csql2)
                            'AlmacenaQuery_sql Csql2, cnn_dbbancos
                            I% = I% + 1
                        End If
                        falta = "0"
                        wnumord = TbDetTmpImp1!F3NUMORD
                    End If
                End If
            End If
            intcont = intcont + 1
            End If
        TbDetTmpImp1.MoveNext
    Loop

    If intcont > 0 Then
        graba = True
    Else
        graba = False
    End If
    
    If TbDetTmpImp1.RecordCount > 0 Then
        TbDetTmpImp1.MoveFirst
        wnumord = TbDetTmpImp1!F3NUMORD
        wtotord = 0#
    Else
        wnumord = 0
        wtotord = 0#
    End If
    
    If Val("" & wnumord) = 0 Then
        wnumord = 0
    End If
    '
    If Not graba Then
        Exit Sub
    End If
    
    'SQL = "Select * from IF3ORDEN WHERE F4NUMORD = " & Val(Format(TbDetTmpImp1!F3NUMORD, "0000000")) & " AND F3CODPRO = '" & Trim(TbDetTmpImp1!F5CODPRO) & "'"
    sql = "Select * from IF3ORDEN WHERE F4NUMORD = '" & TbDetTmpImp1!F3NUMORD & "' AND " _
    & "f3codfab= '" & wf3codfab & "' and f5codmarca='" & wf5codmarca & "'"
    If TbDetOrden1.State = adStateOpen Then TbDetOrden1.Close
    TbDetOrden1.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
    If Not TbDetOrden1.EOF Then
        Do While TbDetOrden1.Fields("F4NUMORD") = Val(Format(wnumord, "0000000"))
            wtotord = wtotord + TOTA
            TbDetOrden1.MoveNext
            If TbDetOrden1.EOF Then Exit Do
        Loop
        '
        sql = "Select * from IF4ORDEN where F4NUMORD=" & Val(Format(wnumord, "0000000")) & ""
        If TbCabOrden1.State = adStateOpen Then TbCabOrden1.Close
        TbCabOrden1.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
        '
        If Not TbCabOrden1.EOF Then
            MONTO1 = Val(Format(wtotord, "0.0000"))
           If falta = "0" And wimporta(I%).f4falta = "0" Then
               FALTA1 = "0"
           Else
               FALTA1 = "1"
           End If
          
           Csql3 = "UPDATE IF4ORDEN SET F4MONTO=" & MONTO1 & ",F4FALTA='" & FALTA1 & "' WHERE F4NUMORD=" & Val(Format(wnumord, "0000000")) & ""
           cnn_dbbancos.Execute (Csql3)
           'AlmacenaQuery_sql Csql3, cnn_dbbancos
           I% = I% + 1
        End If
    End If
        '---------------------ASIGNA DATOS AL DETALLE
        amovs_det(0).campo = "F4NUMIMP": amovs_det(0).valor = "": amovs_det(0).TIPO = "T"
        amovs_det(1).campo = "F3NUMORD": amovs_det(1).valor = "": amovs_det(1).TIPO = "T"
        amovs_det(2).campo = "F3CODFAB": amovs_det(2).valor = "": amovs_det(2).TIPO = "T"
        amovs_det(3).campo = "F5CODPRO": amovs_det(3).valor = "": amovs_det(3).TIPO = "T"
        amovs_det(4).campo = "F3CANTIDAD": amovs_det(4).valor = "": amovs_det(4).TIPO = "N"
        amovs_det(5).campo = "F3PREUNI": amovs_det(5).valor = "": amovs_det(5).TIPO = "N"
        amovs_det(6).campo = "F3TOTAL": amovs_det(6).valor = "": amovs_det(6).TIPO = "N"
        amovs_det(7).campo = "F3PRECOS": amovs_det(7).valor = "": amovs_det(7).TIPO = "N"
        amovs_det(8).campo = "F3MARGEN": amovs_det(8).valor = "": amovs_det(8).TIPO = "N"
        amovs_det(9).campo = "F3VALVTA": amovs_det(9).valor = "": amovs_det(9).TIPO = "N"
        amovs_det(10).campo = "F3DSCTO": amovs_det(10).valor = "": amovs_det(10).TIPO = "N"
        amovs_det(11).campo = "F3VTANET": amovs_det(11).valor = "": amovs_det(11).TIPO = "N"
        amovs_det(12).campo = "f5advalorem": amovs_det(12).valor = "": amovs_det(12).TIPO = "N"
        amovs_det(13).campo = "advalorem": amovs_det(13).valor = "": amovs_det(13).TIPO = "N"
        amovs_det(14).campo = "base": amovs_det(14).valor = "": amovs_det(14).TIPO = "N"
        amovs_det(15).campo = "f2codprov": amovs_det(15).valor = "": amovs_det(15).TIPO = "T"
        amovs_det(16).campo = "f2nomprov": amovs_det(16).valor = "": amovs_det(16).TIPO = "T"
        amovs_det(17).campo = "f5partara": amovs_det(17).valor = "": amovs_det(17).TIPO = "T"
        amovs_det(18).campo = "f5manual": amovs_det(18).valor = "": amovs_det(18).TIPO = "T"
        amovs_det(19).campo = "f3costototal": amovs_det(19).valor = "": amovs_det(19).TIPO = "N"
        amovs_det(20).campo = "f3flete": amovs_det(20).valor = "": amovs_det(20).TIPO = "N"
        amovs_det(21).campo = "f5marca": amovs_det(21).valor = "": amovs_det(21).TIPO = "T"
        amovs_det(22).campo = "f5codmarca": amovs_det(22).valor = "": amovs_det(22).TIPO = "T"
        amovs_det(23).campo = "f3Unimed": amovs_det(23).valor = "": amovs_det(23).TIPO = "T"
        '---------------------Calcula el Numero de Filas
        nitems = 0
        If RSDETALLE.State = adStateOpen Then RSDETALLE.Close
        sql = "Select count(F3ITEM1) as NITEM from TmpDet_Import Where LEN(TRIM(F3ITEM1))> 0 "
    
        RSDETALLE.Open sql, tempo, adOpenDynamic, adLockOptimistic
        If Not RSDETALLE.EOF Then
            nitems = Val("" & RSDETALLE.Fields("NITEM"))
        End If
        RSDETALLE.Close
        ReDim Values(23, nitems)
        If RSDETALLE.State = adStateOpen Then RSDETALLE.Close
        RSDETALLE.Open "Select * from TmpDet_Import", tempo, adOpenStatic, adLockOptimistic
        If Not RSDETALLE.EOF Then
            nfila = 0
            RSDETALLE.MoveFirst
            Do While Not RSDETALLE.EOF
                If (Len(Trim(RSDETALLE.Fields("F3ITEM1") & "")) > 0) And (Len(Trim(RSDETALLE.Fields("F5nompro") & "")) > 0) Then
                    Values(0, nfila) = txtnumero.Text
                    Values(1, nfila) = "" & RSDETALLE.Fields("F3NUMORD")
                    Values(2, nfila) = IIf(Len(Trim(RSDETALLE.Fields("F3CODFAB") & "")) = 0, "", RSDETALLE.Fields("F3CODFAB"))
                    Values(3, nfila) = RSDETALLE.Fields("F5CODPRO")
                    Values(4, nfila) = RSDETALLE.Fields("F3CANTIDAD")
                    Values(5, nfila) = RSDETALLE.Fields("F3PREFOB")
                    Values(6, nfila) = RSDETALLE.Fields("F3TOTAL")
                    Values(7, nfila) = RSDETALLE.Fields("F3PRECOS")
                    Values(8, nfila) = RSDETALLE.Fields("F3MARGEN")
                    Values(9, nfila) = RSDETALLE.Fields("F3VALVTA")
                    Values(10, nfila) = RSDETALLE.Fields("F3DSCTO")
                    Values(11, nfila) = RSDETALLE.Fields("F3VTANET")
                    Values(12, nfila) = RSDETALLE.Fields("f5advalorem")
                    Values(13, nfila) = RSDETALLE.Fields("advalorem")
                    Values(14, nfila) = RSDETALLE.Fields("base")
                    Values(15, nfila) = RSDETALLE.Fields("f2codprov")
                    Values(16, nfila) = RSDETALLE.Fields("f2nomprov")
                    Values(17, nfila) = RSDETALLE.Fields("f5partara")
                    Values(18, nfila) = RSDETALLE.Fields("f5MANUAL")
                    Values(19, nfila) = RSDETALLE.Fields("f3costototal")
                    Values(20, nfila) = RSDETALLE.Fields("f3flete")
                    Values(21, nfila) = RSDETALLE.Fields("f5marca")
                    Values(22, nfila) = RSDETALLE.Fields("f5codmarca")
                    Values(23, nfila) = RSDETALLE.Fields("F5UniMed")
                    nfila = nfila + 1
                End If
                RSDETALLE.MoveNext
            Loop
        End If
        cvalores = "111111111111111111111111"
        cmes = Format(Month(aboFecha.Value), "00")
        If ctipo = "A" Then '---Nuevo
            '-----Graba Cabecera
            GRABA_REGISTRO amovs_cab(), "IMPORT_CAB", ctipo, 24, cnn_dbbancos, ""
            
            If sw_graba_registro = True Then
                '------- GRABA DETALLE '11
                GRABA_REGISTRO_DET amovs_det(), "IMPORT_DET", ctipo, 23, cnn_dbbancos, "", Values(), nfila - 1, cvalores, cmes, ""
            End If
            WNUMERO1 = Val(Format(WNUMERO1, "0000000")) + 1
        Else    '----------Modificacion
            '------- GRABA CABECERA
            GRABA_REGISTRO amovs_cab(), "IMPORT_CAB", ctipo, 22, cnn_dbbancos, "F4NUMIMP = '" & txtnumero.Text & "'"
            
            '------- GRABA DETALLE '11
            
            csql = ("DELETE * FROM IMPORT_DET WHERE F4NUMIMP = '" & txtnumero.Text & "'")
            cnn_dbbancos.Execute csql
            'AlmacenaQuery_sql csql, cnn_dbbancos
            
            GRABA_REGISTRO_DET amovs_det(), "IMPORT_DET", "A", 23, cnn_dbbancos, "F4NUMIMP  = '" & txtnumero.Text & "'", Values(), nfila - 1, cvalores, cmes, ""
        End If
    'End If 'OJO

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
    
    csql = ("DELETE * FROM TB_COSTEODET WHERE F4NUMIMP = '" & txtnumero.Text & "'")
    cnn_dbbancos.Execute csql
    'AlmacenaQuery_sql csql, cnn_dbbancos
    
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
TbDetTmpImp1.Open sql, tempo, adOpenDynamic, adLockOptimistic
'TbDetTmpImp1.MoveFirst
Do While Not TbDetTmpImp1.EOF

    TbDetTmpImp1!F3PRECOS = Format(Val(Format(TbDetTmpImp1!F3PREFOB, "0.0000")) * Val(Format(PnlFactor.Caption, "0.000000000000000")), "###,##0.0000")
    TbDetTmpImp1!f3costototal = Format(TbDetTmpImp1!F3PRECOS * TbDetTmpImp1!F3CANTIDAD, "0.0000")
    'TbDetTmpImp1!F3VALVTA = Format(Val(Format(TbDetTmpImp1!f3PRECOS, "0.0000")) * (1 + Val(Format(TbDetTmpImp1!F3MARGEN, "0.0000")) / 100), "###,##0.0000")
    'TbDetTmpImp1!F3VALVTA = Format(Val(Format(TbDetTmpImp1!f3PRECOS, "0.0000")) * (Val(Format(TbDetTmpImp1!F5FACTOR, "0.0000"))), "###,##0.0000")
    'TbDetTmpImp1!F3VTANET = Format(Val(Format(TbDetTmpImp1!F3VALVTA, "0.0000")) * (1 - Val(Format(TbDetTmpImp1!F3DSCTO, "0.0000")) / 100), "###,##0.0000")
    'TbDetTmpImp1!F3PREUNI = Format(Val(Format(TbDetTmpImp1!F3VTANET, "0.0000")) + (Val(Format(TbDetTmpImp1!F3VTANET, "0.0000")) * 0.19), "###,##0.0000")
    TbDetTmpImp1.UpdateBatch
    'wprecos = wprecos + Val(Format(TbDetTmpImp1!f3PRECOS, "0.0000")) * Val(Format(TbDetTmpImp1!F3CANTIDAD, "0.0000"))
    'wprevta = wprevta + Val(Format(TbDetTmpImp1!F3VALVTA, "0.0000")) * Val(Format(TbDetTmpImp1!F3CANTIDAD, "0.0000"))
    wprecos = wprecos + (TbDetTmpImp1!F3PRECOS * TbDetTmpImp1!F3CANTIDAD)
    wprevta = wprevta + (TbDetTmpImp1!F3VALVTA * TbDetTmpImp1!F3CANTIDAD)
    TbDetTmpImp1.MoveNext
Loop
dxDBGrid2.Dataset.Edit
dxDBGrid2.Dataset.Post
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
    csql = ("DELETE * FROM IMPORT_DET WHERE F4NUMIMP='" & pnumero & "'")
    cnn_dbbancos.Execute csql
    'AlmacenaQuery_sql csql, cnn_dbbancos
    '-------------Eliminaciones
    
    csql = ("DELETE * FROM TB_COSTEOCAB WHERE F4NUMIMP='" & pnumero & "'")
    cnn_dbbancos.Execute csql
    'AlmacenaQuery_sql csql, cnn_dbbancos
    
    csql = ("DELETE * FROM TB_COSTEODET WHERE F4NUMIMP='" & pnumero & "'")
    cnn_dbbancos.Execute csql
    'AlmacenaQuery_sql csql, cnn_dbbancos
    
    llena_items

    sql = "Select * from IF4ORDEN"
    If TbCabOrden1.State = adStateOpen Then TbCabOrden1.Close
    TbCabOrden1.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
    
    sql = "Select * from IF3ORDEN"
    If TbDetOrden1.State = adStateOpen Then TbDetOrden1.Close
    TbDetOrden1.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
    
    sql = "Select * from TmpDet_Import"
    If TbDetTmpImp1.State = adStateOpen Then TbDetTmpImp1.Close
    TbDetTmpImp1.Open sql, tempo, adOpenStatic, adLockOptimistic
    TbDetTmpImp1.MoveFirst
    Do While Not TbDetTmpImp1.EOF
       wnumord = Val(Format(TbDetTmpImp1!F3NUMORD, "0000000"))
        Do While wnumord = Val(Format(TbDetTmpImp1!F3NUMORD, "0000000"))
            wf3codfab = "" & TbDetTmpImp1!F3CODFAB
            wf5codmarca = "" & TbDetTmpImp1!F5CODMARCA
            'SQL = "Select * from IF3ORDEN where F4NUMORD=" & wnumord & " and F3CODPRO='" & TbDetTmpImp1!F5CODPRO & "'"
            sql = "Select * from IF3ORDEN where F4NUMORD='" & wnumord & "' and f3codfab='" & wf3codfab & "' and f5codmarca='" & wf5codmarca & "'"
            If TbDetOrden1.State = adStateOpen Then TbDetOrden1.Close
            TbDetOrden1.Open sql, cnn_dbbancos, adOpenStatic, adLockOptimistic
            If Not TbDetOrden1.EOF Then
               TbDetOrden1.Fields("F3CANFAL") = TbDetOrden1.Fields("F3CANFAL") + Val(Format(TbDetTmpImp1!F3CANTIDAD, "0.000"))
               TbDetOrden1.Update
            End If
            TbDetTmpImp1.MoveNext
            If TbDetTmpImp1.EOF Then
               Exit Do
            End If
        Loop
       TbCabOrden1.Find "F4NUMORD=" & wnumord & "", , adSearchForward
       If Not TbCabOrden1.EOF Then
          TbCabOrden1.Fields("F4FALTA") = "1"
          TbCabOrden1.Fields("F4ESTVAL") = "0"
          TbCabOrden1.Update
       End If
    Loop
End If
End Sub
Private Sub Calcula_Importaciones(pfoco As Integer)
Dim fob         As Double
Dim precos      As Double
Dim ValVta      As Double
Dim vtaneta     As Double
Dim preuni      As Double
Dim margen      As Double
Dim costo       As Double
    With dxDBGrid2
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
                    wprecos = Val(Format(.Columns.ColumnByFieldName("F3PREFOB").Value, "0.0000")) * Val(Format(PnlFactor.Caption, "0.000000000000"))    '8
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
Panel3D2.Enabled = Not Estado
Frame3D1.Enabled = Not Estado
ChkCerrar.Enabled = Not Estado
SSPanel1.Enabled = Not Estado
dxDBGrid1.Enabled = Estado

dxDBGrid2.Columns.ColumnByFieldName("F3NUMORD").DisableEditor = Estado
'dxDBGrid2.Columns.ColumnByFieldName("F3DOCUM").DisableEditor = estado
dxDBGrid2.Columns.ColumnByFieldName("F5CODPRO").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("f2codprov").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("f2nomprov").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("F3CODFAB").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("F5NOMPRO").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("F5UNIMED").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("F3CANTIDAD").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("F3PREFOB").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("f5advalorem").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("advalorem").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("base").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("F3TOTAL").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("F3PRECOS").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("F3MARGEN").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("F3VALVTA").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("F3DSCTO").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("F3VTANET").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("CANTIDAD").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("F3PREUNI").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("f5partara").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("f5manual").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("F3COSTOTOTAL").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("F5MARCA").DisableEditor = Estado
dxDBGrid2.Columns.ColumnByFieldName("F5CODMARCA").DisableEditor = Estado
End Sub

Public Sub costeo()
With rptcosteo
    .datos.ConnectionString = cnn_dbbancos
    sql = "SELECT TB_COSTEODET.*, TB_COSTOSIMP.F2DESCRIPCION " _
    & "FROM TB_COSTEODET INNER JOIN TB_COSTOSIMP ON TB_COSTEODET.F2CODIGO = TB_COSTOSIMP.F2CODIGO " _
    & " where f4numimp='" & txtnumero.Text & "'"
    
    .Caption = "Resumen de Importación"
    .lblempresa.Caption = wempresa
    .lblproforma.Caption = txtnumproforma.Text
    .lblfecha.Caption = aboFecha.Value
    .datos.Source = sql
    .Show vbModal
End With
End Sub
