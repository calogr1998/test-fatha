VERSION 5.00
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Mant_Proveedores 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Proveedores"
   ClientHeight    =   6690
   ClientLeft      =   5115
   ClientTop       =   2520
   ClientWidth     =   9135
   ControlBox      =   0   'False
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
   Icon            =   "Mant_Proveedores.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6690
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5340
      ScaleHeight     =   375
      ScaleWidth      =   2055
      TabIndex        =   45
      Top             =   5880
      Width           =   2055
      Begin CONTROLSLibCtl.dxCheckBox chkOrden 
         Height          =   270
         Left            =   180
         TabIndex        =   10
         Top             =   60
         Width           =   1800
         _Version        =   65536
         _cx             =   3175
         _cy             =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Orden de Compra"
         Enabled         =   -1  'True
         AutoSize        =   -1  'True
         BackStyle       =   0
         BackColor       =   14215660
         ForeColor       =   0
         ViewStyle       =   1
         Checked         =   0   'False
         GroupIndex      =   -1
         TextLayout      =   1
         UseMaskColor    =   -1  'True
         MaskColor       =   12632256
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   635
      ButtonWidth     =   1852
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo "
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar "
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salir     "
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList 
         Left            =   8460
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mant_Proveedores.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mant_Proveedores.frx":05A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mant_Proveedores.frx":0B40
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Mant_Proveedores.frx":10DA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6180
      Left            =   60
      TabIndex        =   26
      Top             =   300
      Width           =   9015
      Begin VB.ComboBox CboCategoria 
         Height          =   330
         ItemData        =   "Mant_Proveedores.frx":1674
         Left            =   6420
         List            =   "Mant_Proveedores.frx":1676
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   5160
         Width           =   2355
      End
      Begin VB.TextBox TxtSiglas 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5820
         MaxLength       =   70
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Txtpostal 
         DataField       =   "F5NOMPRO"
         DataSource      =   "DataProducto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   5
         Top             =   1920
         Width           =   690
      End
      Begin VB.TextBox TxtDesPostal 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1860
         TabIndex        =   34
         Top             =   1920
         Width           =   6975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel 
         Height          =   195
         Index           =   0
         Left            =   7560
         OleObjectBlob   =   "Mant_Proveedores.frx":1678
         TabIndex        =   31
         Top             =   1020
         Width           =   555
      End
      Begin VB.TextBox PnlGrupo 
         Alignment       =   2  'Center
         BackColor       =   &H00EDE9E8&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   30
         Top             =   5580
         Width           =   3015
      End
      Begin VB.TextBox TxtDesCta 
         Alignment       =   2  'Center
         BackColor       =   &H00EDE9E8&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   29
         Top             =   5160
         Width           =   3015
      End
      Begin VB.Frame Frame2 
         Caption         =   "Forma de Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   795
         Left            =   6120
         TabIndex        =   28
         Top             =   4140
         Width           =   2760
         Begin VB.ComboBox cmbfpagos 
            Height          =   330
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   300
            Width           =   2520
         End
      End
      Begin VB.TextBox txtgrupo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   20
         Top             =   5580
         Width           =   705
      End
      Begin VB.Frame Frame7 
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   795
         Left            =   3240
         TabIndex        =   23
         Top             =   4140
         Width           =   2775
         Begin VB.ComboBox CmbTipDoc 
            Height          =   330
            Left            =   180
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   300
            Width           =   2415
         End
      End
      Begin VB.TextBox txtcodcta 
         Alignment       =   2  'Center
         DataField       =   "F5NOMPRO"
         DataSource      =   "DataProducto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   19
         Top             =   5160
         Width           =   690
      End
      Begin VB.TextBox Txtweb 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6000
         MaxLength       =   50
         TabIndex        =   12
         Top             =   3240
         Width           =   2835
      End
      Begin VB.TextBox txtcontacto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   9
         Top             =   2760
         Width           =   7755
      End
      Begin VB.TextBox txtemail 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         TabIndex        =   11
         Top             =   3240
         Width           =   3435
      End
      Begin VB.TextBox txtreferencia 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         MaxLength       =   40
         TabIndex        =   13
         Top             =   3735
         Width           =   7755
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   555
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   3150
         Begin VB.OptionButton opttipo 
            Appearance      =   0  'Flat
            Caption         =   "Extranjero"
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
            Height          =   300
            Index           =   1
            Left            =   1740
            TabIndex        =   25
            Top             =   180
            Width           =   1050
         End
         Begin VB.OptionButton opttipo 
            Appearance      =   0  'Flat
            Caption         =   "Nacional"
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
            Height          =   300
            Index           =   0
            Left            =   420
            TabIndex        =   24
            Top             =   180
            Value           =   -1  'True
            Width           =   960
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   795
         Left            =   120
         TabIndex        =   22
         Top             =   4140
         Width           =   3000
         Begin VB.OptionButton optmoneda 
            Appearance      =   0  'Flat
            Caption         =   "Ambas"
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
            Index           =   2
            Left            =   2040
            TabIndex        =   17
            Top             =   405
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.OptionButton optmoneda 
            Appearance      =   0  'Flat
            Caption         =   "Dólares"
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
            Index           =   1
            Left            =   1080
            TabIndex        =   15
            Top             =   405
            Width           =   900
         End
         Begin VB.OptionButton optmoneda 
            Appearance      =   0  'Flat
            Caption         =   "Soles"
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
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   405
            Width           =   780
         End
      End
      Begin VB.TextBox txtnuevo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         MaxLength       =   11
         TabIndex        =   1
         Top             =   960
         Width           =   1140
      End
      Begin VB.TextBox TxtCuentaContable 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7695
         MaxLength       =   6
         TabIndex        =   8
         Top             =   2340
         Width           =   1140
      End
      Begin VB.TextBox Txtfaxprov 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4260
         MaxLength       =   30
         TabIndex        =   7
         Top             =   2340
         Width           =   1950
      End
      Begin VB.TextBox Txttelprov 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   6
         Top             =   2340
         Width           =   2670
      End
      Begin VB.TextBox Txtdirprov 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1410
         Width           =   7755
      End
      Begin VB.TextBox Txtnomprov 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2295
         MaxLength       =   70
         TabIndex        =   2
         Top             =   960
         Width           =   3435
      End
      Begin VB.TextBox Txtcodprov 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8145
         MaxLength       =   4
         TabIndex        =   27
         Top             =   960
         Width           =   690
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel 
         Height          =   195
         Index           =   1
         Left            =   180
         OleObjectBlob   =   "Mant_Proveedores.frx":16E2
         TabIndex        =   32
         Top             =   1020
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel 
         Height          =   195
         Index           =   2
         Left            =   180
         OleObjectBlob   =   "Mant_Proveedores.frx":174C
         TabIndex        =   33
         Top             =   1500
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel 
         Height          =   195
         Index           =   3
         Left            =   180
         OleObjectBlob   =   "Mant_Proveedores.frx":17BC
         TabIndex        =   35
         Top             =   1980
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel 
         Height          =   195
         Index           =   4
         Left            =   180
         OleObjectBlob   =   "Mant_Proveedores.frx":182E
         TabIndex        =   36
         Top             =   2400
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel 
         Height          =   195
         Index           =   5
         Left            =   180
         OleObjectBlob   =   "Mant_Proveedores.frx":189C
         TabIndex        =   37
         Top             =   2820
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel 
         Height          =   195
         Index           =   6
         Left            =   180
         OleObjectBlob   =   "Mant_Proveedores.frx":190A
         TabIndex        =   38
         Top             =   3300
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel 
         Height          =   195
         Index           =   7
         Left            =   180
         OleObjectBlob   =   "Mant_Proveedores.frx":1972
         TabIndex        =   39
         Top             =   3780
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel 
         Height          =   195
         Index           =   8
         Left            =   3840
         OleObjectBlob   =   "Mant_Proveedores.frx":19E2
         TabIndex        =   40
         Top             =   2400
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel 
         Height          =   195
         Index           =   9
         Left            =   6420
         OleObjectBlob   =   "Mant_Proveedores.frx":1A46
         TabIndex        =   41
         Top             =   2400
         Width           =   1275
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel 
         Height          =   195
         Index           =   10
         Left            =   4920
         OleObjectBlob   =   "Mant_Proveedores.frx":1AC2
         TabIndex        =   42
         Top             =   3300
         Width           =   1035
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel 
         Height          =   195
         Index           =   11
         Left            =   240
         OleObjectBlob   =   "Mant_Proveedores.frx":1B34
         TabIndex        =   43
         Top             =   5220
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel 
         Height          =   195
         Index           =   12
         Left            =   240
         OleObjectBlob   =   "Mant_Proveedores.frx":1BA6
         TabIndex        =   44
         Top             =   5640
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel 
         Height          =   195
         Index           =   13
         Left            =   2340
         OleObjectBlob   =   "Mant_Proveedores.frx":1C20
         TabIndex        =   46
         Top             =   780
         Width           =   1275
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel 
         Height          =   195
         Index           =   14
         Left            =   5880
         OleObjectBlob   =   "Mant_Proveedores.frx":1C96
         TabIndex        =   47
         Top             =   780
         Width           =   1275
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel 
         Height          =   195
         Index           =   15
         Left            =   5460
         OleObjectBlob   =   "Mant_Proveedores.frx":1D00
         TabIndex        =   48
         Top             =   5220
         Width           =   975
      End
   End
End
Attribute VB_Name = "Mant_Proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Af As New ADOFunctions
 
Dim wcodigo         As String * 4
Dim rsproveedores   As New ADODB.Recordset
Dim rst             As New ADODB.Recordset
Dim wgraba          As Integer
Dim wchange As Boolean
Dim wtipdoc As String


Private Sub Actualiza_Proveedor(rucprv As String)
 
    If rsproveedores.State = adStateOpen Then rsproveedores.Close
    If Len(rucprv) = 11 Then
        rsproveedores.Open "SELECT * FROM EF2PROVEEDORES WHERE F2newruc='" & rucprv & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    Else
        rsproveedores.Open "SELECT * FROM EF2PROVEEDORES WHERE F2CODPROV='" & rucprv & "'", cnn_dbbancos, adOpenDynamic, adLockOptimistic
    End If
    
    If Not rsproveedores.EOF Then
        Txtcodprov.Text = "" & rsproveedores.Fields("F2CODPROV")
        Txtnomprov.Text = "" & rsproveedores.Fields("F2NOMPROV")
        TxtSiglas.Text = "" & rsproveedores.Fields("F2NOMabrev")
        txtnuevo.Text = "" & rsproveedores.Fields("F2NEWRUC")
        Txtdirprov.Text = "" & rsproveedores.Fields("F2DIRPROV")
        Txtpostal.Text = "" & rsproveedores.Fields("F7CODPOS")
        Txttelprov.Text = "" & rsproveedores.Fields("F2TELPROV")
        Txtfaxprov.Text = "" & rsproveedores.Fields("F2FAXPROV")
        txtCuentaContable.Text = "" & rsproveedores.Fields("F5CODCTA")
         
        txtEmail.Text = "" & rsproveedores.Fields("F2EMAIL")
        txtContacto.Text = "" & rsproveedores.Fields("F2CONTACTO")
        txtreferencia.Text = "" & rsproveedores.Fields("F2ARTPROV")
        txtWeb.Text = "" & rsproveedores.Fields("F2WEB")
        chkOrden.Checked = IIf(rsproveedores.Fields("f2orden") = True, True, False)
        optmoneda(0).Value = IIf(rsproveedores.Fields("f2tipmon") = "S", True, False)
        optmoneda(1).Value = IIf(rsproveedores.Fields("f2tipmon") = "D", True, False)
        optmoneda(2).Value = IIf(rsproveedores.Fields("f2tipmon") = "A", True, False)
        opttipo(0).Value = IIf(rsproveedores.Fields("F2TIPPROV") = "N", True, False)
        opttipo(1).Value = IIf(rsproveedores.Fields("F2TIPPROV") = "E", True, False)
        
        Call SeleccionaEnComboRight(rsproveedores.Fields("F2TIPDOC") & "", CmbTipDoc)
        Call SeleccionaEnComboRight(rsproveedores.Fields("F2FORPAG") & "", cmbfpagos)
        Call SeleccionaEnComboRight(Format(rsproveedores.Fields("IntCodCategoria") & "", "00000000"), CboCategoria)
        
        txtcodcta.Text = "" & rsproveedores.Fields("F2CODGAS")
        SqlCad = "SELECT NOMBRE FROM BF9GIN WHERE CODIGO= '" & txtcodcta.Text & "' AND BASE='G'"
        If rst.State = adStateOpen Then rst.Close
        rst.Open SqlCad, cnn_dbbancos, adOpenDynamic
        If Not rst.EOF Then
            TxtDesCta.Text = "" & rst.Fields("NOMBRE")
        End If
        rst.Close
        
        SqlCad = "SELECT F2DESZON FROM EF2ZONAS WHERE F2CODZON= '" & Txtpostal.Text & "'"
        If rst.State = adStateOpen Then rst.Close
        rst.Open SqlCad, cnn_dbbancos, adOpenDynamic
        If Not rst.EOF Then
            TxtDesPostal.Text = "" & rst.Fields("F2DESZON")
        End If
        rst.Close
        txtgrupo.Text = "" & rsproveedores.Fields("F2GRUPO")
        SqlCad = "SELECT NOMBRE FROM GRUPOS_FLUJO WHERE CODIGO= '" & txtgrupo.Text & "'"
        If rst.State = adStateOpen Then rst.Close
        rst.Open SqlCad, cnn_dbbancos, adOpenDynamic
        If Not rst.EOF Then
            PnlGrupo.Text = "" & rst.Fields("NOMBRE")
        End If
        rst.Close
        txtnuevo.Enabled = True
        Txtcodprov.Enabled = False
        
    End If
    rsproveedores.Close
    
End Sub

Private Function calcula_codigo()
Dim wnum2   As Integer
Dim SqlCad As String

    SqlCad = "select F2CODPROV from EF2PROVEEDORES  order by F2CODPROV desc "
    If rst.State = adStateOpen Then rst.Close
    rst.Open SqlCad, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not rst.EOF Then
        If rst.Fields("F2CODPROV") & "" = "9999" Then
            rst.MoveNext
            If Not rst.EOF Then
                wnum2 = Val(rst.Fields("F2CODPROV") & "") + 1
            Else
                wnum2 = 1
            End If
        Else
            wnum2 = Val(rst.Fields("F2CODPROV") & "") + 1
        End If
    Else
        wnum2 = 1
    End If
    rst.Close
    wcodigo = Format(wnum2, "0000")
    calcula_codigo = wcodigo
    'TxtCuentaContable.Text = wcodigo

End Function

Private Sub Elimina_Proveedor()

    Beep
    
    If MsgBox("¿Está seguro de eliminar el Proveedor?", vbQuestion + vbYesNo, wnomcia) = vbYes Then
        SqlCad = "SELECT F2CODPROV FROM EF2PROVEEDORES WHERE F2CODPROV= '" & Txtcodprov & "'"
        Set rst = Af.OpenSQLForwardOnly(SqlCad, StrConexDbBancos)
        Actualiza_Log SqlCad, StrConexDbBancos
        If Not rst.EOF Then
            DELETEREC_N "ef2proveedores", StrConexDbBancos, "F2CODPROV = '" & Trim("" & Txtcodprov.Text) & "'"
            Nuevo_Proveedor
        Else
            Beep
        End If
    End If
    
End Sub

Private Sub chkorden_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtgrupo.SetFocus
    End If

End Sub

Private Sub CboCategoria_Click()
If right(CboCategoria.Text, 8) = "99999999" Then
        '    MsgBox "Nuevo"
        Mant_Categorias.Show 1
        Unload Mant_Categorias
        Set Mant_Categorias = Nothing
        CargaCategoria
        Call SeleccionaEnComboRight(Format(wcodgrupo, "00000000"), CboCategoria)
End If
End Sub

Private Sub Form_Load()
 
If cnn_dbbancos.State = 1 Then cnn_dbbancos.Close
cnn_dbbancos.Open StrConexDbBancos

    'Me.Left = MDIBancos.dxsidebar.Width
    'Me.Top = 1050
    CargaCategoria
    CargaDocumentos
    CargaForPag
    
    If sw_nuevo_documento = True Then
        Nuevo_Proveedor
        wgraba = 1
    Else
        Actualiza_Proveedor Cod_Prove
    End If
    wchange = False
    
    
End Sub

Private Sub CargaCategoria()
    Dim RsCat As New ADODB.Recordset
    
    csql = "select IntCodCategoria,StrDesCategoria from Categoria order by intcodcategoria"
    
    Set RsCat = Af.OpenSQLForwardOnly(csql, StrConexDbBancos)
    
    CboCategoria.Clear
    
    If RsCat.RecordCount > 0 Then
        RsCat.MoveFirst
        
        Do While Not RsCat.EOF
            CboCategoria.AddItem RsCat!STRDESCATEGORIA & Space(299) & Format(RsCat!INTCODCATEGORIA, "00000000")
            RsCat.MoveNext
        Loop
        
            CboCategoria.AddItem "Nueva categoría" & Space(299) & "99999999"
    End If
End Sub

Private Sub CargaForPag()
Dim tbfpagos1 As New ADODB.Recordset
    SqlCad = "Select * from ef2forpag order by f2despag"
    If tbfpagos1.State = adStateOpen Then tbfpagos1.Close
    tbfpagos1.Open SqlCad, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not tbfpagos1.EOF Then
        Do While Not tbfpagos1.EOF
            If Len(Trim(tbfpagos1.Fields("f2tipo") & "")) > 0 Then
                If tbfpagos1.Fields("f2tipo") & "" = "F" Or tbfpagos1.Fields("f2tipo") & "" = "C" Then
                    cmbfpagos.AddItem tbfpagos1.Fields("f2despag") & Space(100) & tbfpagos1.Fields("f2forpag") & ""
                End If
            End If
            tbfpagos1.MoveNext
        Loop
    End If
    tbfpagos1.Close
    cmbfpagos.ListIndex = 0
End Sub

Private Sub CargaDocumentos()
Dim TbDocumento1 As New ADODB.Recordset
    SqlCad = "Select * from documentos"
    If TbDocumento1.State = adStateOpen Then TbDocumento1.Close
    TbDocumento1.Open SqlCad, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If Not TbDocumento1.EOF Then
        Do While Not TbDocumento1.EOF
            CmbTipDoc.AddItem UCase(TbDocumento1.Fields("F2DESDOC")) + Space(100) + TbDocumento1.Fields("F2CODDOC")
            TbDocumento1.MoveNext
        Loop
    End If
    CmbTipDoc.ListIndex = 0
    TbDocumento1.Close
    
End Sub

Private Sub Graba_Proveedor()
On Error GoTo graba
Dim ctipo  As String
    
    Dim amovs(0 To 20) As a_grabacion
    
    wcodigo = "" & Trim(Txtcodprov.Text)
    SqlCad = "select * from ef2proveedores where f2codprov = '" & wcodigo & "'"
    If rst.State = adStateOpen Then rst.Close
    rst.Open SqlCad, cnn_dbbancos, adOpenDynamic, adLockOptimistic
    If rst.EOF Then
        ctipo = "A"
    Else
        ctipo = "M"
    End If
    rst.Close
    amovs(0).campo = "F2CODPROV": amovs(0).valor = wcodigo: amovs(0).Tipo = "T"
    amovs(1).campo = "F2NOMPROV": amovs(1).valor = Trim("" & Txtnomprov.Text): amovs(1).Tipo = "T"
    amovs(2).campo = "F2NEWRUC": amovs(2).valor = Trim("" & txtnuevo.Text): amovs(2).Tipo = "T"
    amovs(3).campo = "F2DIRPROV": amovs(3).valor = Trim("" & Txtdirprov.Text): amovs(3).Tipo = "T"
    amovs(4).campo = "F7CODPOS": amovs(4).valor = Trim("" & Txtpostal.Text): amovs(4).Tipo = "T"
    amovs(5).campo = "F2TELPROV": amovs(5).valor = Trim("" & Txttelprov.Text): amovs(5).Tipo = "T"
    amovs(6).campo = "F2FAXPROV": amovs(6).valor = Trim("" & Txtfaxprov.Text): amovs(6).Tipo = "T"
    amovs(7).campo = "F5CODCTA": amovs(7).valor = Trim("" & txtCuentaContable.Text): amovs(7).Tipo = "T"
    amovs(8).campo = "F2ARTPROV": amovs(8).valor = Trim("" & txtreferencia.Text): amovs(8).Tipo = "T"
    Rem NSE amovs(9).campo = "F2ORDEN": amovs(9).valor = IIf(chkOrden.checked = True, "Si", "No"): amovs(9).TIPO = "T"
    amovs(9).campo = "F2TIPMON"
    If optmoneda(0).Value = True Then
        amovs(9).valor = "S": amovs(9).Tipo = "T"
    End If
    If optmoneda(1).Value = True Then
        amovs(9).valor = "D": amovs(9).Tipo = "T"
    End If
    If optmoneda(2).Value = True Then
        amovs(9).valor = "A": amovs(9).Tipo = "T"
    End If
    amovs(10).campo = "F2TIPPROV"
    If opttipo(0).Value = True Then
        amovs(10).valor = "N": amovs(10).Tipo = "T"
    End If
    If opttipo(1).Value = True Then
        amovs(10).valor = "E": amovs(10).Tipo = "T"
    End If
    amovs(11).campo = "F2EMAIL": amovs(11).valor = Trim("" & txtEmail.Text): amovs(11).Tipo = "T"
    amovs(12).campo = "F2CONTACTO": amovs(12).valor = Trim("" & txtContacto.Text): amovs(12).Tipo = "T"
    amovs(13).campo = "F2WEB": amovs(13).valor = Trim("" & txtWeb.Text): amovs(13).Tipo = "T"
    amovs(14).campo = "F2CODGAS": amovs(14).valor = Trim("" & txtcodcta.Text): amovs(14).Tipo = "T"
    amovs(15).campo = "F2TIPDOC": amovs(15).valor = right(Trim("" & CmbTipDoc.Text), 2): amovs(15).Tipo = "T"
    amovs(16).campo = "F2ORDEN": amovs(16).valor = IIf(chkOrden.Checked, "1", "0"): amovs(16).Tipo = "T"
    amovs(17).campo = "F2GRUPO": amovs(17).valor = Trim(txtgrupo.Text): amovs(17).Tipo = "T"
    amovs(18).campo = "F2FORPAG": amovs(18).valor = right(Trim("" & cmbfpagos.Text), 3): amovs(18).Tipo = "T"
    amovs(19).campo = "F2nomabrev": amovs(19).valor = TxtSiglas.Text: amovs(19).Tipo = "T"
    amovs(20).campo = "intCodCategoria": amovs(20).valor = IIf(CboCategoria.ListIndex > -1, Val(right(CboCategoria.Text, 8)), 0): amovs(20).Tipo = "N"
    
    txtnuevo.Enabled = True
    GRABA_REGISTRO amovs(), "EF2PROVEEDORES", ctipo, 20, StrConexDbBancos, "F2CODPROV = '" & wcodigo & "'"
        
    If ctipo = "A" Then
        MsgBox "El proveedor ha sido registrado", vbInformation, wnomcia
    Else
        MsgBox "El proveedor ha sido actualizado", vbInformation, wnomcia
    End If
    wgraba = 0
    sw_nuevo_documento = False
    txtnuevo.Enabled = True
    Txtcodprov.Enabled = False
    
    
    Exit Sub
    
graba:
    If Err = 3186 Then
        MsgBox "La base de Datos esta Bloqueada por otro Usuario espere unos segundos...", 48, "CONTROL Plus!"
        Resume
    Else
        MsgBox "Se ha producido el sgte. error " & Error(Err), 48, "CONTROL Plus!"
        
        Resume Next
    End If
        
End Sub

Public Function AddProveedorMySQL(ByVal oRuc As String, ByVal oRazonsocial As String, ByVal oContraseña As String) As Boolean

    On Error GoTo ErrorHandler
    MsgBox " Ruc: " & oRuc & " - Nombre : " & oRazonsocial & " - Contraseña:" & oContraseña
    AjouterCleint = False
    
    Dim Rs As New Recordset
    Set Rs = New Recordset
    Dim conn As ADODB.Connection
    Dim strIPAddress As String
    
    Set CON = New ADODB.Connection
        CON.CursorLocation = adUseClient
        CON.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
        & "SERVER=108.167.182.78;" _
        & " DATABASE=calogr19_10081962991;" _
        & "UID=calogr19_richi;PWD=richi123; OPTION=3"
        CON.Open
    Set Rs = Nothing
        Rs.CursorLocation = adUseClient
        SqlCad = "INSERT INTO usuarios (usuario_dni,usuario_nombre,usuario_contraseña,celulares_id,tipo_usuario_id) VALUES ('" _
        & oRuc & "','" & oRazonsocial & "','" & oContraseña & "','8','3')"

    Rs.Open SqlCad, CON, 3, 3

    Set Rs = Nothing

    CON.Close

    Set CON = Nothing

    AddProveedorMySQL = True
ErrorHandler:
    MsgBox Err.Number & vbLf & Err.Description & vbLf & Err.HelpContext & vbLf & Err.Source, , ""
End Function
 
Private Sub Nuevo_Proveedor()
    
    txtnuevo.Enabled = True
    Txtcodprov.Enabled = True
    txtnuevo.TabIndex = 0
    txtnuevo.Enabled = True
    Txtcodprov.Enabled = True
    Txtcodprov.Text = calcula_codigo
    Txtnomprov.Text = ""
    TxtSiglas.Text = ""
    txtnuevo.Text = ""
    Txtdirprov.Text = ""
    Txtpostal.Text = ""
    Txttelprov.Text = ""
    Txtfaxprov.Text = ""
    TxtDesPostal.Text = ""
    txtCuentaContable.Text = ""
    txtEmail.Text = ""
    txtContacto.Text = ""
    txtreferencia.Text = ""
    txtcodcta.Text = ""
    TxtDesCta.Text = ""
    txtgrupo.Text = ""
    PnlGrupo.Text = ""
    chkOrden.Checked = False
    
    optmoneda(0).Value = True
    optmoneda(1).Value = False
    optmoneda(2).Value = False
    opttipo(0).Value = True
    opttipo(1).Value = False
    txtnuevo.TabIndex = 0
    wnuevoproveedor = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'If cnn_dbbancos.State = 1 Then cnn_dbbancos.Close
'Set cnn_DbBancos = Nothing
End Sub

Private Sub optmoneda_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        If optmoneda(0).Value Then
            optmoneda(1).SetFocus
        Else
            If optmoneda(1).Value Then
                optmoneda(2).SetFocus
            Else
                If optmoneda(2).Value Then
                    opttipo(0).SetFocus
                End If
            End If
        End If
    End If

End Sub

Private Sub opttipo_Click(Index As Integer)
On Error Resume Next
    Dim RsPrv As New ADODB.Recordset
    If opttipo(1).Value Then
        If Len(Trim(Cod_Prove)) = 0 Then
        SqlCad = "select f2newruc from ef2proveedores where val(f2newruc) < 10000000000 order by f2newruc desc"
        If RsPrv.State = adStateOpen Then RsPrv.Close
        RsPrv.Open SqlCad, cnn_dbbancos, 3, 1
        If Not RsPrv.EOF Then
            txtnuevo.Text = Format(Val(RsPrv.Fields(0).Value) + 1, "00000000000")
        Else
            txtnuevo.Text = "00000000001"
        End If
        RsPrv.Close
        End If
        Txtnomprov.SetFocus
    Else
        txtnuevo.Text = ""
        txtnuevo.SetFocus
    End If
    
End Sub

Private Sub opttipo_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 And opttipo(0).Value Then
        opttipo(1).Value = True
        
    End If
    If KeyAscii = 13 And opttipo(1).Value Then
        Txtnomprov.SetFocus
    End If

End Sub





Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Trim(Button.Caption)
        Case "Nuevo"
            Me.MousePointer = vbHourglass
            
            If wgraba = 1 And wchange Then
                If MsgBox("No ha guardado los cambios, ¿desea hacerlo ahora", vbYesNo + vbInformation, "CONTROL Plus!") = 6 Then
                    Graba_Proveedor
                End If
            End If
            
            Nuevo_Proveedor
            
            txtnuevo.SetFocus
            sw_nuevo_documento = True
            wgraba = 1
            wchange = False
            
            Me.MousePointer = vbDefault
        Case "Grabar"
            Me.MousePointer = vbHourglass
            
            If CboCategoria.ListIndex = -1 Then
                MsgBox "Debe seleccionar una categoría", vbExclamation, wnomcia
                CboCategoria.SetFocus
                Exit Sub
            End If
            
            SqlCad = "select * from ef2proveedores where  F2NEWRUC='" & Trim(txtnuevo.Text) & "'"
            
            If rsproveedores.State = adStateOpen Then rsproveedores.Close
            
            rsproveedores.Open SqlCad, cnn_dbbancos, adOpenDynamic, adLockOptimistic
            
            If rsproveedores.EOF Then
                Graba_Proveedor
            Else
                If sw_nuevo_documento = True Then
                    MsgBox "El Proveedor con el RUC " & txtnuevo.Text & " existe, Ingrese otro", vbInformation, "CONTROL Plus!"
                Else
                    Graba_Proveedor
                End If
            End If
            
            If rsproveedores.State = 1 Then rsproveedores.Close
            
            Me.MousePointer = vbDefault
            
            wcodcliprov = Me.Txtcodprov.Text
            
            Unload Me
            
            If FrmName = "LISTA_PROVEEDORES" Then
                Lista_Proveedores.Show
            End If
        Case "Eliminar"
            Me.MousePointer = vbHourglass
            Elimina_Proveedor
            Me.MousePointer = vbDefault
        Case "Salir"
            
            Me.Hide
            
            If FrmName = "LISTA_PROVEEDORES" Then
                Lista_Proveedores.Show
            End If
    End Select
End Sub

Private Sub txtcodcta_Change()
If sw_nuevo_documento = False Then wchange = True
TxtDesCta.Text = ""
End Sub

Private Sub txtcodcta_DblClick()
txtcodcta_KeyDown 113, 0
End Sub

Private Sub txtcodcta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
'        llampro = 1
'        txtcodcta.SetFocus
        wdestino = "E"
        ayuda_gastos.TipoConcepto = "E"
        ayuda_gastos.Show 1
        Unload ayuda_gastos
        txtcodcta.Text = wgastos
        txtcodcta_KeyPress 13
    End If
End Sub

Private Sub txtcodcta_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim tbcomtab1 As ADODB.Recordset
    Set tbcomtab1 = New ADODB.Recordset
    If KeyAscii = 13 Then
        gcodppp = txtcodcta.Text
        
        SqlCad = "Select * from bf9gin where codigo='" & gcodppp & "' and base='G'"
        If tbcomtab1.State = adStateOpen Then tbcomtab1.Close
        tbcomtab1.Open SqlCad, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not tbcomtab1.EOF Then
'           gcueppp = tbcomtab1.Fields("cuenta") & ""
'           txtcodcta.Text = gcueppp
           TxtDesCta.Text = tbcomtab1.Fields("nombre").Value & ""
           
        Else
            txtcodcta.Text = ""
            TxtDesCta.Text = ""
            MsgBox "El Codigo ingresado no existe. Vuelva a Ingresarlo ", vbInformation, wnomcia
                           
            txtcodcta.SetFocus
        End If
        tbcomtab1.Close

    End If
End Sub

Private Sub txtcodprov_KeyPress(KeyAscii As Integer)
On Error GoTo ERRORCLI
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Txtcodprov.Text = Trim("" & Txtcodprov.Text)
        SqlCad = "select * from ef2proveedores where f2codprov = '" & Txtcodprov.Text & "'"
        If rsproveedores.State = adStateOpen Then rsproveedores.Close
        rsproveedores.Open SqlCad, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsproveedores.EOF Then
            Actualiza_Proveedor Txtcodprov.Text
        Else
            Txtnomprov.SetFocus
        End If
        rsproveedores.Close
    End If
    Exit Sub
    
ERRORCLI:
    Resume Next
        
End Sub

Private Sub txtcontacto_Change()
If sw_nuevo_documento = False Then wchange = True
End Sub

Private Sub txtcontacto_GotFocus()
    
    txtContacto.SelStart = 0: txtContacto.SelLength = Len(txtContacto.Text)

End Sub

Private Sub txtContacto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtEmail.SetFocus
    End If

End Sub

Private Sub TxtCuentaContable_DblClick()
TxtCuentaContable_KeyDown 113, 0
End Sub

Private Sub TxtCuentaContable_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
    wctacont = ""
    Ayuda_PlanCta.Show 1
    Unload Ayuda_PlanCta
    
    DoEvents
    If Len(Trim(wctacont)) > 0 Then
        txtCuentaContable.Text = Trim("" & wctacont)
    End If
End If
End Sub

Private Sub Txtdirprov_Change()
If sw_nuevo_documento = False Then wchange = True
End Sub

Private Sub Txtdirprov_GotFocus()
    
    Txtdirprov.SelStart = 0: Txtdirprov.SelLength = Len(Txtdirprov.Text)
    
End Sub

Private Sub TxtDirprov_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then Txtpostal.SetFocus
    
End Sub

Private Sub txtemail_Change()
If sw_nuevo_documento = False Then wchange = True
End Sub

Private Sub txtemail_GotFocus()
    
    If Len(Trim(txtEmail.Text)) = 0 Then
        txtEmail.Text = "@"
    End If
    txtEmail.SelStart = 0: txtEmail.SelLength = Len(txtEmail.Text)

End Sub

Private Sub txtemail_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtreferencia.SetFocus
    End If

End Sub

Private Sub Txtfaxprov_Change()
If sw_nuevo_documento = False Then wchange = True
End Sub

Private Sub Txtfaxprov_GotFocus()
    
    Txtfaxprov.SelStart = 0: Txtfaxprov.SelLength = Len(Txtfaxprov.Text)
    
End Sub

Private Sub txtgrupo_DblClick()
    txtgrupo_KeyDown 113, 0
End Sub

Private Sub txtgrupo_GotFocus()
    txtgrupo.SelStart = 0: txtgrupo.SelLength = Len(txtgrupo.Text)
End Sub

Private Sub txtgrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        sw_ayuda = True
        wdestino = "P"
        wcodgrupo = ""
        ayuda_grupoflujo.Show 1
        sw_ayuda = False
        If Len(Trim(wcodgrupo)) > 0 Then
            txtgrupo.Text = wcodgrupo
            PnlGrupo.Text = wdesgrupo
        End If
        txtgrupo_KeyPress 13
    End If
End Sub

Private Sub txtgrupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
optmoneda(0).SetFocus
End If

End Sub

Private Sub txtgrupo_LostFocus()
Dim RSCONSULTA      As New ADODB.Recordset

    If sw_ayuda = False Then
        If Len(Trim(txtgrupo.Text)) > 0 Then
            strSQL = "SELECT NOMBRE FROM GRUPOS_FLUJO WHERE CODIGO='" & txtgrupo.Text & "' AND LEFT(CODIGO,1) = 'P'"
            
            If RSCONSULTA.State = adStateOpen Then RSCONSULTA.Close
            RSCONSULTA.Open strSQL, cnn_dbbancos, adOpenStatic, adLockOptimistic
            If Not RSCONSULTA.EOF Then
                PnlGrupo.Text = "" & RSCONSULTA.Fields(0)
            Else
                PnlGrupo.Text = "": txtgrupo.Text = ""
                MsgBox "Código del Grupo de Flujo no existe. Verifique.", vbInformation, "CONTROL Plus!"
                txtgrupo.SetFocus
            End If
            RSCONSULTA.Close
            Set RSCONSULTA = Nothing
        Else
                PnlGrupo.Text = "": txtgrupo.Text = ""
        End If
    End If
End Sub

Private Sub Txtnomprov_Change()
If sw_nuevo_documento = False Then wchange = True

End Sub

Private Sub Txtnomprov_GotFocus()
    
    Txtnomprov.SelStart = 0: Txtnomprov.SelLength = Len(Txtnomprov.Text)
    
End Sub

Private Sub Txtnomprov_KeyPress(KeyAscii As Integer)
    
    'KeyAscii = TxtNum1(KeyAscii)
    If KeyAscii = 13 Then
        TxtSiglas.SetFocus
        
    End If

End Sub

Public Function TxtNum1(KeyAscii As Integer) As Integer

    If (KeyAscii < 48 Or KeyAscii > 57) Then
        TxtNum1 = KeyAscii
    Else
        TxtNum1 = 0
        Beep
    End If

End Function

Private Sub Txtfaxprov_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtContacto.SetFocus
    End If
    KeyAscii = TxtNum(KeyAscii)
    
End Sub

Private Sub txtnuevo_LostFocus()
txtnuevo.Text = Format(txtnuevo.Text, "00000000000")
End Sub

Private Sub txtpostal_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 Then
        wcodzona = ""
        'hlp_zonas.Show 1
        ayuda_zonas.Show 1
        If Len(wcodzona) > 0 Then
            Txtpostal.Text = wcodzona
            txtpostal_KeyPress 13
        End If
    End If
    
End Sub

Private Sub txtreferencia_Change()
If sw_nuevo_documento = False Then wchange = True
End Sub

Private Sub TxtSiglas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Txtdirprov.SetFocus
End If
End Sub

Private Sub Txttelprov_Change()
If sw_nuevo_documento = False Then wchange = True
End Sub

Private Sub Txttelprov_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Txtfaxprov.SetFocus
    End If
    KeyAscii = TxtNum(KeyAscii)
    
End Sub

Public Function TxtNum(KeyAscii As Integer) As Integer

    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 32 Then
         TxtNum = KeyAscii
    Else
         TxtNum = 0
         Beep
    End If

End Function

Private Sub txtnuevo_KeyPress(KeyAscii As Integer)
Dim SqlCad As String
    
    If KeyAscii = 13 Then
        If opttipo.ITEM(0).Value = True Then
        txtnuevo.Text = Trim("" & txtnuevo.Text)
            If Len(txtnuevo.Text) <> 11 Then
                  MsgBox "El RUC tiene que tener 11 dígitos", vbInformation, "CONTROL Plus!"
                  txtnuevo.SetFocus
                  Exit Sub
            'ElseIf Verifica_Ruc(txtnuevo.Text) = False Then
            '    MsgBox "El RUC no es válido. Verifique", vbInformation, "CONTROL Plus!"
            '      txtnuevo.SetFocus
            '      Exit Sub
            End If
        End If
        SqlCad = "select * from ef2proveedores where f2newruc = '" & txtnuevo.Text & "'"
        If rsproveedores.State = adStateOpen Then rsproveedores.Close
        rsproveedores.Open SqlCad, cnn_dbbancos, adOpenDynamic, adLockOptimistic
        If Not rsproveedores.EOF Then
            MsgBox "El RUC ya existe.", vbInformation, "CONTROL Plus!"
            Actualiza_Proveedor rsproveedores.Fields("f2newruc")
        Else
            valida_sunat txtnuevo.Text
            Txtdirprov = ruc_direccion
            Txtnomprov.Text = ruc_rsocial
            If Len(Trim(ruc_telefono)) = 7 And IsNumeric(ruc_telefono) Then
                Txttelprov.Text = ruc_telefono
            End If
            Txtnomprov.SetFocus
        End If
        If rsproveedores.State = adStateOpen Then rsproveedores.Close
        Txtnomprov.SetFocus
    Else
        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtreferencia_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        chkOrden.SetFocus
    End If
    KeyAscii = TxtNum1(KeyAscii)

End Sub

Private Sub Txttelprov_GotFocus()
    
    Txttelprov.SelStart = 0: Txttelprov.SelLength = Len(Txttelprov.Text)
    
End Sub

Private Sub txtpostal_DblClick()
    
    txtpostal_KeyDown 113, 0
    
End Sub

Private Sub txtpostal_GotFocus()

    Txtpostal.SelStart = 0: Txtpostal.SelLength = Len(Txtpostal.Text)
    
End Sub

Private Sub Txtpostal_Change()
If sw_nuevo_documento = False Then wchange = True
    If Len(Txtpostal.Text) = 0 Then
        TxtDesPostal.Text = ""
    End If

End Sub

Private Sub txtpostal_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If Len(Trim(Txtpostal.Text)) > 0 Then
            SqlCad = "SELECT F2DESZON FROM EF2ZONAS WHERE F2CODZON= '" & Txtpostal.Text & "'"
            Set rst = Af.OpenSQLForwardOnly(SqlCad, StrConexDbBancos)
            If Not rst.EOF Then
                TxtDesPostal.Text = "" & rst.Fields("F2DESZON")
                Txttelprov.SetFocus
            Else
                MsgBox "El Codigo no Existe. Ingrese un nuevo Codigo", vbInformation, "CONTROL Plus!"
                Txtpostal.Text = ""
                Txtpostal.SetFocus
                
            End If
            rst.Close
        Else
            MsgBox "Debe Ingresar un Codigo de Distrito", vbInformation, "CONTROL Plus!"
            Txtpostal.SetFocus
        End If
    End If

End Sub

Private Sub Txtweb_Change()
If sw_nuevo_documento = False Then wchange = True
End Sub
