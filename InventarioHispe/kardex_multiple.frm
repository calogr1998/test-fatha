VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{03F7CB5F-9E40-4B74-A3ED-7DBEAAB01C6C}#1.0#0"; "aBox.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form kardex_MULTIPLE 
   Caption         =   "Kardex"
   ClientHeight    =   7035
   ClientLeft      =   1140
   ClientTop       =   855
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   10230
   Begin MSAdodcLib.Adodc adoproductos 
      Height          =   375
      Left            =   120
      Top             =   6600
      Visible         =   0   'False
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoproductos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2445
      Left            =   120
      TabIndex        =   5
      Top             =   -30
      Width           =   10065
      _Version        =   65536
      _ExtentX        =   17754
      _ExtentY        =   4313
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
      Begin Threed.SSFrame SSFrame3 
         Height          =   840
         Left            =   180
         TabIndex        =   10
         Top             =   1440
         Width           =   5670
         _Version        =   65536
         _ExtentX        =   10001
         _ExtentY        =   1482
         _StockProps     =   14
         Caption         =   " Moneda "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSOption optmoneda 
            Height          =   240
            Index           =   0
            Left            =   990
            TabIndex        =   11
            Top             =   330
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Soles"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optmoneda 
            Height          =   240
            Index           =   1
            Left            =   3480
            TabIndex        =   12
            Top             =   330
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   423
            _StockProps     =   78
            Caption         =   "Dólares"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSCheck chkalmacen 
         Height          =   240
         Left            =   6120
         TabIndex        =   2
         Top             =   225
         Width           =   2850
         _Version        =   65536
         _ExtentX        =   5027
         _ExtentY        =   423
         _StockProps     =   78
         Caption         =   "Todos los Almacenes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   885
         Left            =   180
         TabIndex        =   6
         Top             =   360
         Width           =   5670
         _Version        =   65536
         _ExtentX        =   10001
         _ExtentY        =   1561
         _StockProps     =   14
         Caption         =   " Rango de Fechas "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin aBoxCtl.aBox abodesde 
            Height          =   315
            Left            =   945
            TabIndex        =   0
            Top             =   330
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            ABoxType        =   ""
            MinValue        =   "D01000101"
            MaxValue        =   "D99991231"
            ABoxStyle       =   2
            Alignment       =   1
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
            ApplyTextFormat =   -1  'True
            TextFormat      =   "dd/mm/yyyy"
            Text            =   "31/01/2007"
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
            ButtonPicture   =   "kardex_MULTIPLE.frx":0000
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
         Begin aBoxCtl.aBox abohasta 
            Height          =   315
            Left            =   3990
            TabIndex        =   1
            Top             =   330
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            ABoxType        =   ""
            MinValue        =   "D01000101"
            MaxValue        =   "D99991231"
            ABoxStyle       =   2
            Alignment       =   1
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
            ApplyTextFormat =   -1  'True
            TextFormat      =   "dd/mm/yyyy"
            Text            =   "31/01/2007"
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
            ButtonPicture   =   "kardex_MULTIPLE.frx":0352
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3405
            TabIndex        =   8
            Top             =   375
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   270
            TabIndex        =   7
            Top             =   375
            Width           =   465
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid vsfalmacenes 
         Height          =   1785
         Left            =   6120
         TabIndex        =   3
         ToolTipText     =   "Presione F2 para editar columna"
         Top             =   540
         Width           =   3795
         _cx             =   6694
         _cy             =   3149
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   0
         BackColorSel    =   8388608
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   2
         GridLinesFixed  =   3
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"kardex_MULTIPLE.frx":06A4
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   1
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.ListBox lstselec 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   7200
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   2835
      Width           =   2955
   End
   Begin ActiveToolBars.SSActiveToolBars SSActiveToolBars1 
      Left            =   4275
      Top             =   6600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131082
      ToolBarsCount   =   1
      ToolsCount      =   3
      Tools           =   "kardex_MULTIPLE.frx":0709
      ToolBars        =   "kardex_MULTIPLE.frx":2D2B
   End
   Begin TrueOleDBGrid70.TDBGrid tdbproductos 
      Height          =   3645
      Left            =   45
      TabIndex        =   13
      Top             =   2835
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   6429
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   "F5CODPRO"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Modelo"
      Columns(1).DataField=   "F5CODFAB"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Descripción"
      Columns(2).DataField=   "F5NOMPRO"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Marca"
      Columns(3).DataField=   "F2DESMAR"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   609
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1667"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1588"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=2249"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2170"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=4736"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=4657"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(21)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   -2147483633
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=42"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=42"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=36"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=3"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(46)  =   "Named:id=33:Normal"
      _StyleDefs(47)  =   ":id=33,.parent=0"
      _StyleDefs(48)  =   "Named:id=34:Heading"
      _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   ":id=34,.wraptext=-1"
      _StyleDefs(51)  =   "Named:id=35:Footing"
      _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(53)  =   "Named:id=36:Selected"
      _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=37:Caption"
      _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(57)  =   "Named:id=38:HighlightRow"
      _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=39:EvenRow"
      _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(61)  =   "Named:id=40:OddRow"
      _StyleDefs(62)  =   ":id=40,.parent=33"
      _StyleDefs(63)  =   "Named:id=41:RecordSelector"
      _StyleDefs(64)  =   ":id=41,.parent=34"
      _StyleDefs(65)  =   "Named:id=42:FilterBar"
      _StyleDefs(66)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Lista de Productos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   90
      TabIndex        =   14
      Top             =   2535
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Productos Seleccionados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7680
      TabIndex        =   9
      Top             =   2535
      Width           =   2100
   End
End
Attribute VB_Name = "kardex_MULTIPLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs  As New ADODB.Recordset
Dim col As TrueOleDBGrid70.Column
Dim cols As TrueOleDBGrid70.Columns

Private Function getFilter() As String
Dim cadena As String
Dim n As Integer
    
    For Each col In cols
        If Trim(col.FilterText) <> "" Then
            n = n + 1
            If n > 1 Then
                cadena = cadena & " AND "
            End If
            cadena = cadena & col.DataField & " LIKE '" & col.FilterText & "*'"
        End If
    Next col
    getFilter = cadena

End Function

Private Sub abodesde_GotFocus()

    abodesde.FocusSelect = True

End Sub

Private Sub abodesde_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        abohasta.SetFocus
    End If

End Sub

Private Sub abohasta_GotFocus()
    
    abohasta.FocusSelect = True
    
End Sub

Private Sub abohasta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        chkalmacen.SetFocus
    End If

End Sub

Private Sub chkalmacen_Click(Value As Integer)

     If Value = 0 Then
        vsfalmacenes.Enabled = True
        LLENA_PRODUCTOS_ALMACEN
    Else
        vsfalmacenes.Enabled = False
        LLENA_PRODUCTOS
    End If

End Sub

Private Sub Form_Activate()

    If lstselec.ListCount = 0 Then
        SSActiveToolBars1.Tools(1).Enabled = False
    Else
        SSActiveToolBars1.Tools(1).Enabled = True
    End If
    
End Sub

Private Sub Form_Load()

    Me.Height = 7550
    
    Me.Width = 10530
    Me.Left = 1500
    Me.Top = 1050
    
    Me.AutoRedraw = False

    abodesde.Value = Format(Date, "dd/mm/yyyy")
    abohasta.Value = Format(Date, "dd/mm/yyyy")

    LLENA_ALMACENES
    If chkalmacen.Value Then
        vsfalmacenes.Enabled = False
        LLENA_PRODUCTOS
    Else
        vsfalmacenes.Enabled = True
        LLENA_PRODUCTOS_ALMACEN
    End If
    
    Me.AutoRedraw = True

End Sub

Private Sub lstselec_DblClick()

    If lstselec.ListIndex >= 0 Then
        lstselec.RemoveItem lstselec.ListIndex
    End If
    
    If lstselec.ListCount = 0 Then
        SSActiveToolBars1.Tools(1).Enabled = False
    Else
        SSActiveToolBars1.Tools(1).Enabled = True
    End If
    
End Sub


Private Sub SSActiveToolBars1_ToolClick(ByVal Tool As ActiveToolBars.SSTool)

    Select Case Tool.Id
        Case "ID_Imprimir":
            If lstselec.ListCount > 0 Then
                imprimir
            Else
                MsgBox "Falta seleccionar el producto.", vbInformation, "Atención"
                lstselec.SetFocus
            End If
        Case "ID_Pasarproductos"
            Pasar_productos
        Case "ID_Salir":
            Unload Me
    End Select
    
End Sub

Private Sub LLENA_PRODUCTOS()
Dim csql        As String
Dim ncont       As Integer
Dim I           As Integer
Dim cad         As String

    ncont = 0
    'csql = "SELECT IF5PLA.F5CODPRO AS F5CODPRO, IF5PLA.F5NOMPRO AS F5NOMPRO,  IF5PLA.F5CODFAB AS F5CODFAB FROM IF5PLA"
    'csql = csql & " ORDER BY IF5PLA.F5NOMPRO"
    
    cad = ""
    csql = "SELECT IF5PLA.F5CODPRO, IF5PLA.F5NOMPRO, IF5PLA.F5CODFAB, EF2MARCAS.F2DESMAR " & _
           "FROM IF5PLA LEFT JOIN EF2MARCAS ON IF5PLA.F5MARCA = EF2MARCAS.F2CODMAR " & _
           "ORDER BY IF5PLA.F5NOMPRO;"
    
    adoproductos.ConnectionString = cnn_dbbancos
    adoproductos.RecordSource = csql
    adoproductos.Refresh
    Set tdbproductos.DataSource = adoproductos.Recordset
        
End Sub

Private Sub LLENA_ALMACENES()
Dim Row     As Long

    rsalmacen.Open "SELECT F2CODALM,F2NOMALM FROM EF2ALMACENES ORDER BY F2CODALM", cnn_dbbancos
    If Not rsalmacen.EOF Then
        vsfalmacenes.Rows = 1
        Row = 1
        Do While Not rsalmacen.EOF
            vsfalmacenes.AddItem vbTab & rsalmacen.Fields("F2CODALM") & vbTab & rsalmacen.Fields("F2NOMALM"), Row
            rsalmacen.MoveNext
            Row = Row + 1
        Loop
    End If
    rsalmacen.Close

End Sub

Private Sub tdbproductos_DblClick()
Dim nlong   As Integer

    nlong = Len(Trim(tdbproductos.Columns(0)))
    lstselec.AddItem Trim(tdbproductos.Columns(0)) & Space(10 - nlong) & " - " & Space(3) & Trim(tdbproductos.Columns(1)) & " - " & Space(3) & Trim(tdbproductos.Columns(2))
    SSActiveToolBars1.Tools(1).Enabled = True
    
End Sub

Private Sub tdbproductos_FilterChange()
On Error GoTo errhandler
Set cols = tdbproductos.Columns
Dim c As Integer

    c = tdbproductos.col
    tdbproductos.HoldFields
    adoproductos.Recordset.Filter = getFilter()
    tdbproductos.col = c
    tdbproductos.EditActive = True
    Exit Sub
    
errhandler:
    
    MsgBox Err.Source & ":" & vbCrLf & Err.Description
    For Each col In tdbproductos.Columns
        col.FilterText = ""
    Next col
    adoproductos.Recordset.Filter = adFilterNone

End Sub


Private Sub vsfalmacenes_Click()

    Me.MousePointer = 11
    lstselec.Clear
    LLENA_PRODUCTOS_ALMACEN
    Me.MousePointer = 1
    
End Sub

Private Sub imprimir()
Dim csql        As String
Dim ccodprod    As String
Dim cad         As String
Dim I           As Integer

    'Construye lista de productos seleccionados
    cad = "": wcod_alm = ""
    For I = 0 To lstselec.ListCount - 1
        cad = cad & "'" & Trim(Left$(lstselec.List(I), 11)) & "',"
    Next I
    cad = Mid$(cad, 1, Len(cad) - 1)
    
    acr_kardex.fldempresa.Text = wnomcia
    acr_kardex.fldtitulo.Text = "Del  " & Format(abodesde.Value, "dd/mm/yyyy") & "  al  " & Format(abohasta.Value, "dd/mm/yyyy")
    acr_kardex.fldfecha.Text = Format(Date, "dd/mm/yyyy")
    
    If chkalmacen.Value = False Then
        vsfalmacenes.col = 1
        wcod_alm = vsfalmacenes.Text
    End If
    
    If Len(wcod_alm) > 0 Then
        acr_kardex.lblalmacen.Visible = True
        acr_kardex.fldcodalmacen.Visible = True
        acr_kardex.fldnomalmacen.Visible = True
        vsfalmacenes.col = 1: acr_kardex.fldcodalmacen.Text = vsfalmacenes.Text
        vsfalmacenes.col = 2: acr_kardex.fldnomalmacen.Text = vsfalmacenes.Text
    Else
        acr_kardex.lblalmacen.Visible = False
        acr_kardex.fldcodalmacen.Visible = False
        acr_kardex.fldnomalmacen.Visible = False
    End If
    
    acr_kardex.datconexion.ConnectionString = cconex_dbbancos
    If optmoneda(0).Value = True Then
        If Len(Trim(wcod_alm)) = 0 Then
            csql = "SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO, IF4VALES.F4NUMVAL, F7CODMED, " & _
               "IF4VALES.F4TIPDOC,IF4VALES.F4SERDOC,IF4VALES.F4NUMDOC, " & _
               "IF3VALES.F3VALVTA, IF4VALES.F4FECVAL, IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,0) AS ENTRADAK, " & _
               "IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3CANPRO,0) AS SALIDAK, " & _
               "IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3VALVTA ,0) AS ENTRADACOS, " & _
               "IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3VALVTA,0) AS SALIDACOS, IF3VALES.F2CODALM,SF1ORIGENES.F1NOMORI " & _
               "FROM (IF4VALES INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) " & _
               "ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) " & _
               "INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI " & _
               "Where ((IF5PLA.F5CODPRO) in (" & cad & ")) And ((IF4VALES.F4FECVAL) >=cvdate('" & abodesde.Value & "') " & _
               "And (IF4VALES.F4FECVAL) <= cvdate('" & abohasta.Value & "'))  " & _
               "ORDER BY IF5PLA.F5CODPRO,IF4VALES.F4FECVAL,IF4VALES.F4NUMVAL;"
        Else
            csql = "SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO, IF4VALES.F4NUMVAL, F7CODMED, " & _
               "IF4VALES.F4TIPDOC,IF4VALES.F4SERDOC,IF4VALES.F4NUMDOC, IF3VALES.F3VALVTA, IF4VALES.F4FECVAL, " & _
               "IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,0) AS ENTRADAK, " & _
               "IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3CANPRO,0) AS SALIDAK, " & _
               "IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3VALVTA ,0) AS ENTRADACOS, " & _
               "IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3VALVTA,0) AS SALIDACOS, IF3VALES.F2CODALM,SF1ORIGENES.F1NOMORI " & _
               "FROM (IF4VALES INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) " & _
               "ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) " & _
               "INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI " & _
               "Where ((IF5PLA.F5CODPRO) in (" & cad & ")) And ((IF4VALES.F4FECVAL) >=cvdate('" & abodesde.Value & "') " & _
               "And (IF4VALES.F4FECVAL) <= cvdate('" & abohasta.Value & "')) And ((IF3VALES.F2CODALM) = '" & wcod_alm & "') " & _
               "ORDER BY IF5PLA.F5CODPRO,IF4VALES.F4FECVAL,IF4VALES.F4NUMVAL;"
        End If
        acr_kardex.LblIngreso.Caption = "S/."
        acr_kardex.LblSalida.Caption = "S/."
        acr_kardex.LblSaldos.Caption = "S/."
    Else
        If Len(Trim(wcod_alm)) = 0 Then
            csql = "SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO, IF4VALES.F4NUMVAL, F7CODMED, " & _
               "IF4VALES.F4TIPDOC,IF4VALES.F4SERDOC,IF4VALES.F4NUMDOC, IF3VALES.F3VALDOL, IF4VALES.F4FECVAL, " & _
               "IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,0) AS ENTRADAK, " & _
               "IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3CANPRO,0) AS SALIDAK, " & _
               "IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3VALDOL ,0) AS ENTRADACOS, " & _
               "IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3VALDOL,0) AS SALIDACOS, IF3VALES.F2CODALM,SF1ORIGENES.F1NOMORI " & _
               "FROM (IF4VALES INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) " & _
               "ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) " & _
               "INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI " & _
               "Where ((IF5PLA.F5CODPRO) in (" & cad & ")) And ((IF4VALES.F4FECVAL) >=cvdate('" & abodesde.Value & "') And (IF4VALES.F4FECVAL) <= cvdate('" & abohasta.Value & "')) " & _
               "ORDER BY IF5PLA.F5CODPRO,IF4VALES.F4FECVAL, IF4VALES.F4NUMVAL;"
        Else
            csql = "SELECT IF5PLA.F5CODPRO, IF5PLA.F5CODFAB, IF5PLA.F5NOMPRO, IF4VALES.F4NUMVAL, F7CODMED, " & _
               "IF4VALES.F4TIPDOC,IF4VALES.F4SERDOC,IF4VALES.F4NUMDOC, IF3VALES.F3VALDOL, IF4VALES.F4FECVAL, " & _
               "IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3CANPRO,0) AS ENTRADAK, " & _
               "IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3CANPRO,0) AS SALIDAK, " & _
               "IIf(Left(IF3VALES.F4NUMVAL,1)='I',IF3VALES.F3VALDOL ,0) AS ENTRADACOS, " & _
               "IIf(Left(IF3VALES.F4NUMVAL,1)='S',IF3VALES.F3VALDOL,0) AS SALIDACOS, IF3VALES.F2CODALM,SF1ORIGENES.F1NOMORI " & _
               "FROM (IF4VALES INNER JOIN (IF5PLA INNER JOIN IF3VALES ON IF5PLA.F5CODPRO = IF3VALES.F5CODPRO) " & _
               "ON (IF4VALES.F4NUMVAL = IF3VALES.F4NUMVAL) AND (IF4VALES.F2CODALM = IF3VALES.F2CODALM)) " & _
               "INNER JOIN SF1ORIGENES ON IF4VALES.F1CODORI = SF1ORIGENES.F1CODORI " & _
               "Where ((IF5PLA.F5CODPRO) in (" & cad & ")) And ((IF4VALES.F4FECVAL) >=cvdate('" & abodesde.Value & "') And (IF4VALES.F4FECVAL) <= cvdate('" & abohasta.Value & "')) And ((IF3VALES.F2CODALM) = '" & wcod_alm & "') " & _
               "ORDER BY IF5PLA.F5CODPRO,IF4VALES.F4FECVAL,IF4VALES.F4NUMVAL;"
        End If
                
        acr_kardex.LblIngreso.Caption = "US$"
        acr_kardex.LblSalida.Caption = "US$"
        acr_kardex.LblSaldos.Caption = "US$"
    End If
    
    acr_kardex.datconexion.Source = csql
    acr_kardex.Show vbModal
    
End Sub

Private Sub LLENA_PRODUCTOS_ALMACEN()
Dim csql        As String
Dim ncont       As Integer
Dim I           As Integer
Dim var         As String

    wcod_alm = vsfalmacenes.TextMatrix(vsfalmacenes.Row, 1)
    ncont = 1
    csql = "SELECT A.F5CODPRO AS F5CODPRO, A.F5NOMPRO AS F5NOMPRO,A.F5CODFAB AS F5CODFAB, EF2MARCAS.F2DESMAR FROM EF2MARCAS INNER JOIN (IF6ALMA AS B INNER JOIN IF5PLA AS A ON B.F5CODPRO = A.F5CODPRO) ON EF2MARCAS.F2CODMAR = A.F5MARCA WHERE (B.F2CODALM)='" & wcod_alm & "' "
    csql = csql & " ORDER BY A.F5NOMPRO"
    
    adoproductos.ConnectionString = cnn_dbbancos
    adoproductos.RecordSource = csql
    adoproductos.Refresh
    
    Set tdbproductos.DataSource = adoproductos.Recordset

End Sub
Private Sub Pasar_productos()
Dim nlong   As Integer
Dim x As Integer

    With adoproductos.Recordset
    .MoveFirst
    For x = 1 To .RecordCount
        nlong = Len(Trim(tdbproductos.Columns(0)))
        lstselec.AddItem Trim(tdbproductos.Columns(0)) & Space(10 - nlong) & " - " & Space(3) & Trim(tdbproductos.Columns(1)) & " - " & Space(3) & Trim(tdbproductos.Columns(2))
        .MoveNext
    Next x
        SSActiveToolBars1.Tools(1).Enabled = True
        
    End With
End Sub
