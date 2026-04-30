VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frm_Iã‚çsìoò^ 
   Caption         =   "ã‚çsìoò^"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14085
   Icon            =   "frm_Iã‚çsìoò^.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   14085
   StartUpPosition =   3  'Windows ÇÃä˘íËíl
   Begin VB.CommandButton ï€ë∂ 
      Caption         =   "ï€ë∂ÅiF11)"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   29
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton ï¬Ç∂ÇÈ 
      Caption         =   "ï¬Ç∂ÇÈ(F12)"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12240
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton B_SET 
      Caption         =   "SET"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   10320
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Caption         =   "ìoò^"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   13815
      Begin VB.CommandButton çÌèú 
         Caption         =   "ÉfÅ[É^çÌèú"
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3120
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox ã‚çsî‘çÜ 
         Height          =   330
         IMEMode         =   3  'µÃå≈íË
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox ã‚çsñº 
         Height          =   330
         IMEMode         =   4  'ëSäpÇ–ÇÁÇ™Ç»
         Left            =   2040
         TabIndex        =   7
         Top             =   720
         Width           =   4215
      End
      Begin éÿä∑ÇΩÇÎÇ§.ZU020_ComboBox éxï•ì˙ 
         Height          =   345
         Left            =   2040
         TabIndex        =   13
         Top             =   1080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   609
         ForeColor       =   -2147483640
         ForeColor       =   -2147483640
         IMEMode         =   3
         TextWidth       =   615
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         P8_ListBoxMax   =   0
         P8_KeySort      =   0   'False
      End
      Begin éÿä∑ÇΩÇÎÇ§.ZU020_ComboBox âcã∆ì˙ 
         Height          =   345
         Left            =   2040
         TabIndex        =   14
         Top             =   1440
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   609
         ForeColor       =   -2147483640
         ForeColor       =   -2147483640
         IMEMode         =   3
         TextWidth       =   615
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         P8_ListBoxMax   =   0
         P8_KeySort      =   0   'False
      End
      Begin éÿä∑ÇΩÇÎÇ§.ZU020_ComboBox óòëßãÊï™ 
         Height          =   345
         Left            =   2040
         TabIndex        =   17
         Top             =   1800
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   609
         ForeColor       =   -2147483640
         ForeColor       =   -2147483640
         IMEMode         =   3
         TextWidth       =   615
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         P8_ListBoxMax   =   0
         P8_KeySort      =   0   'False
      End
      Begin éÿä∑ÇΩÇÎÇ§.ZU020_ComboBox óòëßì˙êî 
         Height          =   345
         Left            =   2040
         TabIndex        =   19
         Top             =   2160
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   609
         ForeColor       =   -2147483640
         ForeColor       =   -2147483640
         IMEMode         =   3
         TextWidth       =   615
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         P8_ListBoxMax   =   0
         P8_KeySort      =   0   'False
      End
      Begin éÿä∑ÇΩÇÎÇ§.ZU020_ComboBox óòëßéxï• 
         Height          =   345
         Left            =   9240
         TabIndex        =   21
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   609
         ForeColor       =   -2147483640
         ForeColor       =   -2147483640
         IMEMode         =   3
         TextWidth       =   615
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         P8_ListBoxMax   =   0
         P8_KeySort      =   0   'False
      End
      Begin éÿä∑ÇΩÇÎÇ§.ZU020_ComboBox óòëßçTèú 
         Height          =   345
         Left            =   9240
         TabIndex        =   23
         Top             =   720
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   609
         ForeColor       =   -2147483640
         ForeColor       =   -2147483640
         IMEMode         =   3
         TextWidth       =   615
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         P8_ListBoxMax   =   0
         P8_KeySort      =   0   'False
      End
      Begin éÿä∑ÇΩÇÎÇ§.ZU020_ComboBox ã‡óòåvéZ 
         Height          =   345
         Left            =   9240
         TabIndex        =   25
         Top             =   1080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   609
         ForeColor       =   -2147483640
         ForeColor       =   -2147483640
         IMEMode         =   3
         TextWidth       =   615
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         P8_ListBoxMax   =   0
         P8_KeySort      =   0   'False
      End
      Begin VB.Label Label3 
         BackColor       =   &H00D6DBBD&
         Caption         =   " ã‡óòåvéZì˙êî"
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7320
         TabIndex        =   26
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00D6DBBD&
         Caption         =   " óòëßçTèúãÊï™"
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7320
         TabIndex        =   24
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackColor       =   &H00D6DBBD&
         Caption         =   " óòëßéxï•ï˚ñ@"
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7320
         TabIndex        =   22
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00D6DBBD&
         Caption         =   " óòëßåvéZì˙êî"
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H00D6DBBD&
         Caption         =   " óòëßãÊï™"
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00D6DBBD&
         Caption         =   " éxï•ì˙"
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00D6DBBD&
         Caption         =   " âcã∆ì˙"
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D6DBBD&
         Caption         =   " ã‚çsñº"
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00D6DBBD&
         Caption         =   " ã‚çsî‘çÜ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label L_ã‚çsñº 
         BackColor       =   &H00D6DBBD&
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         TabIndex        =   9
         Top             =   720
         Width           =   4215
      End
   End
   Begin VB.CommandButton Command4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   9120
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   9120
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   9120
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   120
      Top             =   9960
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   635
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3885
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   6853
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1041
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin éÿä∑ÇΩÇÎÇ§.ZU070_Label êVãKïœçX 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BackColor_Shape1=   8454016
      BackColor_Shape2=   8421504
      BorderColor_Shape1=   49152
      BorderColor_Shape2=   4210752
      ForeColor       =   255
      Caption         =   "êVãKïœçX"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label ÉÅÉbÉZÅ[ÉW 
      BackColor       =   &H00C0C000&
      Caption         =   "ÉÅÉbÉZÅ[ÉW"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   8640
      Width           =   15015
   End
End
Attribute VB_Name = "frm_Iã‚çsìoò^"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
