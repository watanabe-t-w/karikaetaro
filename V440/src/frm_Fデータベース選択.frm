VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_Fデータベース選択 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "データベース選択"
   ClientHeight    =   10230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13155
   Icon            =   "frm_Fデータベース選択.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   13155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton 最新GRID 
      Caption         =   "最新グリッド表示(&G)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   120
      TabIndex        =   17
      Top             =   6480
      Width           =   12855
      Begin VB.TextBox 企業名Key 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   4  '全角ひらがな
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   24
         Top             =   960
         Width           =   2895
      End
      Begin VB.CommandButton 保存 
         Caption         =   "新規企業登録"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10200
         TabIndex        =   23
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton 金剛石処理 
         Caption         =   "金剛石の処理"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10200
         TabIndex        =   22
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CheckBox 削除 
         Alignment       =   1  '右揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2040
         Width           =   255
      End
      Begin VB.TextBox 企業名 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   4  '全角ひらがな
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   20
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CheckBox 実績 
         Alignment       =   1  '右揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2040
         Width           =   255
      End
      Begin VB.ComboBox 備考 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   4  '全角ひらがな
         Left            =   1800
         TabIndex        =   18
         Text            =   "備考"
         Top             =   1680
         Width           =   2895
      End
      Begin 借換たろう.ZU070_Label 新規変更 
         Height          =   495
         Left            =   240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BackColor_Shape1=   8454016
         BackColor_Shape2=   8421504
         BorderColor_Shape1=   49152
         BorderColor_Shape2=   4210752
         ForeColor       =   255
         Caption         =   "企業名登録"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 最新処理日"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5040
         TabIndex        =   40
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 保存日､時刻"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5040
         TabIndex        =   37
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 復元日"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5040
         TabIndex        =   36
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 削除日"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5040
         TabIndex        =   35
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 作成日"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5040
         TabIndex        =   33
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 備考"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   39
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 企業名Key"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   38
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 削除"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   34
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 企業名"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   32
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label L_作成日 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6600
         TabIndex        =   31
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label L_削除日 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6600
         TabIndex        =   30
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label L_復元日 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6600
         TabIndex        =   29
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label L_保存日時刻 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6600
         TabIndex        =   28
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label L_最新処理日 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6600
         TabIndex        =   27
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label L_実績 
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 実績共有"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   26
         Top             =   2040
         Width           =   1575
      End
   End
   Begin VB.CommandButton 運用管理 
      Caption         =   "運用管理"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9600
      Width           =   1575
   End
   Begin VB.CommandButton B_2 
      Caption         =   "本支店管理"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9600
      Width           =   1575
   End
   Begin VB.CommandButton B_3 
      Caption         =   "復元処理"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9600
      Width           =   1575
   End
   Begin VB.CommandButton B_4 
      Caption         =   "DB最適化"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9600
      Width           =   1575
   End
   Begin VB.CommandButton B_5 
      Caption         =   "バックアップ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9600
      Width           =   1575
   End
   Begin VB.CommandButton B_6 
      Caption         =   "ログ一覧表"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9600
      Width           =   1575
   End
   Begin VB.CommandButton 閉じる 
      Caption         =   "終了(F12)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'なし
      Height          =   840
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   12735
      Begin VB.CommandButton 検索 
         Caption         =   "検索(&S)"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10800
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox 検索備考 
         Height          =   300
         IMEMode         =   4  '全角ひらがな
         Left            =   7560
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0FFFF&
         Caption         =   "備考"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6960
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label L_SerDirName 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   5
         Top             =   430
         Width           =   12135
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ﾏｲｺﾝﾋﾟｭﾀｰ名"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label L_MyComName 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1920
         TabIndex        =   3
         Top             =   120
         Width           =   4695
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   10680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   0
      Top             =   10680
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
         Name            =   "ＭＳ ゴシック"
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
      Height          =   4125
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   7276
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   14
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
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
   Begin 借換たろう.ZU050_Button ラベル_帳票名 
      Height          =   615
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "データベース選択"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   20.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label メッセージ 
      BackColor       =   &H00C0C000&
      Caption         =   "メッセージ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3360
      TabIndex        =   16
      Top             =   4920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   800
      Left            =   240
      Top             =   960
      Width           =   12705
   End
End
Attribute VB_Name = "frm_Fデータベース選択"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'
Private Const pPROGRAM_ID As String = "FAA020_ログインユーザ選択"
'
'------------------------------------------------
' 修正履歴
'------------------------------------------------
' @001 2018/05/30 CSV出力先フォルダ指定
'
Dim wRs2 As ADODB.Recordset
Dim wstr As String

Dim CopyFLG As Integer    '複製判断フラグ　false:新規　true:複製
Dim CopyDB名  As String   '複製対象DB名

Dim FLG_KIGYO As Boolean, FLG_DIR As Boolean, FLG_LOSTK As Boolean, FLG_New As Boolean
Dim wi_Kosin1 As Integer, wi_Kosin2 As Integer
Dim wl_st As Long
Dim wDB名 As String, wBK As String
Dim w企業名Key As String, wFlg企業名Key As String, wNew企業名Key As String

Dim wi決算月 As Integer, wi決算締日 As Integer, wi回収有無 As Integer, wi支払有無 As Integer, wd実績年月 As Date

'----------< Msg >------------------------------------------------------------------
Private Const Msg_01 = "しばらくしてから処理を行ってください"
'
'------------------------------------------------
' Form_Initialize
'------------------------------------------------
Private Sub Form_Initialize()
'
    FLG_DIR = False
    wBK = Left$(GCurDir + "\" + "backup", 200)
    
    '----------< LOG WRITE >--------------------------------------------------------
    GStr = "0," & "9," & P8.FCChr(GMyComputerName, 16) & "," & P8.FCChr("", 30) & "," & P8.FCChr("", 30)
    GRet = PUT_LOG_FILE(GStr)
'
End Sub

'------------------------------------------------
' Form_Load
'------------------------------------------------
Private Sub Form_Load()
'
    Dim wDb As New ADODB.Connection
    Dim wRs3 As ADODB.Recordset
    
    Dim ws01 As String, ws02 As String
'
    On Error GoTo Form_Load_ERR
'
    ' =========================================
    '                 初期設定
    ' =========================================
    Me.Caption = GVerNo
    
    DataGrid1.Visible = True
    
    企業名Key.MaxLength = 50
    企業名.MaxLength = 50
'
    FLG_KIGYO = False
    wi_Kosin1 = 0: wi_Kosin2 = 0
    
    L_SerDirName.Caption = Left$(GSerDir, 50)
    L_MyComName.Caption = Left$(GMyComputerName, 16)
    'L_CurDirName.Caption = GCurDir
    'L_SerComName.Caption = GSerComputerName
    
    GKeyName = ""

    If GSys.Sys = "LUFU" Then
        実績.Visible = False
        L_実績.Visible = False
    End If

    If GSys.Ker <> True And GSys.Han <> True Then
        B_2.Caption = ""
        B_2.Enabled = False
    End If

    '金剛石 or 借換たろう！
'    金剛石処理.Caption = GProduct & "の処理"
    金剛石処理.Caption = "実行"
    If GProduct <> "金剛石" Then
        実績.Visible = False
        L_実績.Visible = False
        B_2.Caption = ""
        B_2.Enabled = False
    End If
'
    If FLG_DIR = False Then
    On Error Resume Next
        wBK = Left$(GCurDir + "\" + "backup", 200)
    
        '----------< List.mdb Open >------------------------------------------------
        Call AdoDbOpen("Jet", wDb, GCurDir + "\" + GMain, "", , GPwd)
    
        ' =========================================
        '           フォルダ位置
        ' =========================================
        wstr = ""
        wstr = wstr + "Select フォルダ位置 "
        wstr = wstr + "From LIST000_データ保存先マスタ "
        wstr = wstr + "Where System = 'System'"
        Call AdoRecordsetOpen(wDb, wRs3, wstr)
        If Not wRs3.eof Then
            ws01 = P8.FCStr(wRs3("フォルダ位置"))
            If ws01 = "" Or ws01 = "C:\金剛石" Or ws01 = "C:\借換たろう" Then
                wRs3("フォルダ位置") = wBK
                wRs3.Update
                
                ws02 = Dir(wBK, vbDirectory)
                If ws02 = "" Then
                    MkDir wBK
                    Err.Clear
                End If
            Else
                ws02 = Dir(ws01, vbDirectory)
                If ws02 = "" Then
                    ws02 = Dir(wBK, vbDirectory)
                    If ws02 = "" Then
                        MkDir wBK
                        Err.Clear
                    End If
                    
                    wRs3("フォルダ位置") = wBK
                    wRs3.Update
                Else
                    wBK = ws01
                End If
            End If
        
        End If
        wRs3.Close
        Set wRs3 = Nothing
        
        '----------< List.mdb Close >-----------------------------------------------
        wDb.Close
        Set wDb = Nothing
        
        FLG_DIR = True
    End If
'
    '----------< List.mdb Open >----------------------------------------------------
    Call AdoDbOpen("Jet", GDb2, GSerDir + "\" + GMain, "", , GPwd)
    
    ' =========================================
    '           初期フラグ解除
    ' =========================================
    Call FLG_OFFMYPC
    
    Call BUTTON_TOOLTIPTEXT_SET
'
    Call 登録後初期セット
    メッセージ = ""
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Form_Load_ERR:
    pERR_MES = pPROGRAM_ID + "/ Form_Load() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' Form_Activate
'------------------------------------------------
Private Sub Form_Activate()
'
    Dim wDb As New ADODB.Connection
    Dim wSDate As String
'
    On Error GoTo Form_Activate_ERR
'
    '*** 画面のちらつきをなくす為の Doevents
    DoEvents
'
    GKeyName = ""
    GDbName = ""
    If FLG_KIGYO = True Then
        
        '----------< List.mdb Open >------------------------------------------------
        Call AdoDbOpen("Jet", GDb2, GSerDir + "\" + GMain, "", , GPwd)
        
        DoEvents
        '
        ' =========================================
        '           KIGYOSHORI_END
        ' =========================================
        wSDate = Format(Now, "yyyy/mm/dd hh:nn:ss")
        Call KIGYOSHORI_END(企業名Key, wSDate)
        
        '----------< KXXX.mdb Open >------------------------------------------------
        Call AdoDbOpen("Jet", wDb, GSerDir + "\" + wDB名, "", , GPwd)
    
        '----------< KXXX.mdb >-----------------------------------------------------
        '----------< DAAA070_企業名マスタ >-----------------------------------------
        wstr = "Update"
        wstr = wstr + " DAAA070_企業名マスタ"
        wstr = wstr + " Set"
        wstr = wstr + " 処理終了日付 = #" + wSDate + "#,"
        wstr = wstr + " 保存日 = #" + wSDate + "#,"
        wstr = wstr + " 入力中端末名 = '',"
        wstr = wstr + " 処理中端末名 = ''"
        wstr = wstr + " Where 企業名Key = '" & 企業名Key & "'"
        wDb.Execute wstr
    
        '----------< KXXX.mdb Close >-----------------------------------------------
        wDb.Close
        Set wDb = Nothing
        
        '----------< LOG WRITE >----------------------------------------------------
        GStr = "3," & "9," & P8.FCChr(GMyComputerName, 16) & "," & P8.FCChr(企業名Key, 30) & "," & P8.FCChr(企業名, 30)
        GRet = PUT_LOG_FILE(GStr)
        
        '
        Adodc1.Recordset.Open
        DataGrid1.Visible = True
        '
        
        Call 画面セット(False)
        
        FLG_KIGYO = False
        
        '
        DoEvents
        
        '----------< BUTTON_ENABLE_SET >--------------------------------------------
        Call BUTTON_ENABLE_SET(True)
        '
        DoEvents
        
    End If
'
    Call CEkey.AllSelect
    Call CEkey.SetFs(企業名Key, True)
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Form_Activate_ERR:
    pERR_MES = pPROGRAM_ID + "/ Form_Activate() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' Form_KeyDown
'------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        Call 閉じる_Click
    End If
End Sub

'------------------------------------------------
' Form_KeyPress
'------------------------------------------------
Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = CEkey.X020_EnterKey(Me, KeyAscii, True)
    メッセージ = ""
End Sub

'------------------------------------------------
' AdodcRefresh
'------------------------------------------------
Private Sub AdodcRefresh()
'
    On Error GoTo AdodcRefresh_ERR
'
    ' =========================================
    '             グリッドの初期値
    ' =========================================
    Call MXA030_DataGridInit(DataGrid1)
 
    Set DataGrid1.DataSource = Adodc1
    
    ' =========================================
    '              ConnectionString
    ' =========================================
    Call AdodcSet(Adodc1, GDb2)
    
    ' =========================================
    '              メインクエリ
    ' =========================================
    GWhere = ""
    If 検索備考.Text <> "" Then
        GWhere = GWhere & " And 備考 = '" + 検索備考.Text + "'"
    End If
    GWhere = " Where (1=1) " + GWhere
    
    wstr = ""
    wstr = wstr + "Select "
    
    If GSys.Sit = True Then
        '1
        wstr = wstr + " IIf(企業区分='連結親会社','1',IIf(企業区分='連結本部','2',IIf(企業区分='連結子会社','3',"
        wstr = wstr + "IIf(企業区分='全社','4',IIf(企業区分='本部','5',IIf(企業区分='支店','6','9')))))),"
    End If
    
    wstr = wstr + "企業名Key,企業名,備考,削除日,実績データ共有,入力中端末名,処理中端末名,端末コンピュータ名,"
    wstr = wstr + "親会社名,企業区分,支店コード,支店名,"
    wstr = wstr + "処理開始日付,処理終了日付,最新処理日,作成日,保存日,復元日,"
    wstr = wstr + "企業名Key As Grd企業名Key,"
    wstr = wstr + "企業名 As Grd企業名,"
    wstr = wstr + "備考 As Grd備考,"
    wstr = wstr + "IIF(isnull(削除日),'','×') As Grd削除,"
    '
    If GSys.Sys = "FULL" Then
        wstr = wstr + "IIF(実績データ共有=1,'○','×') As Grd実績共有,"
    End If
    If GSys.Lan = True Then
        wstr = wstr + "IIF(入力中端末名<>'',入力中端末名,'') As Grd入力中,"
        wstr = wstr + "IIF(処理中端末名<>'',処理中端末名,'') As Grd処理中,"
        wstr = wstr + "端末コンピュータ名 As Grd前回使用, "
    End If
    '
    If GSys.Sit = True Then
        wstr = wstr + "IIF(親会社名<>'',親会社名,IIF(企業区分='連結親会社',企業区分,"
        wstr = wstr + " IIF(企業区分='全社' and 親会社名='',支店コード,''))) As Grd親会社名,"
        wstr = wstr + "企業区分 As Grd企業区分, "
        wstr = wstr + "IIF(支店コード<>'',支店コード,'単独') As Grd支店コード,"
        wstr = wstr + "支店名 As Grd支店名, "
    End If
    '
    wstr = wstr + "Format$(処理開始日付,'yyyy/mm/dd hh:nn:ss') As Grd処理開始日付,"
    wstr = wstr + "Format$(処理終了日付,'yyyy/mm/dd hh:nn:ss') As Grd処理終了日付,"
    '
    wstr = wstr + "Format$(最新処理日,'yyyy/mm/dd') As Grd最新処理日,"
    wstr = wstr + "Format$(作成日,'yyyy/mm/dd') As Grd作成日,"
    wstr = wstr + "Format$(保存日,'yyyy/mm/dd hh:nn:ss') As Grd保存日,"
    wstr = wstr + "Format$(復元日,'yyyy/mm/dd hh:nn:ss') As Grd復元日"
    wstr = wstr + " From DAAA070_企業名マスタ "
    wstr = wstr + GWhere
    
    If GSys.Sit = True Then
        wstr = wstr + "Order by 削除日,1,支店コード,企業名Key,企業名"
    Else
        wstr = wstr + "Order by 削除日,企業名Key,企業名"
    End If
  
   Adodc1.RecordSource = wstr
    Adodc1.Refresh
    
    Call XZMA010_DataGrid_Init
        Call XZMA010_DataGrid_Set("企業名Key", "", 1200, "L")
        Call XZMA010_DataGrid_Set("企業名", "", 3000, "L")
        Call XZMA010_DataGrid_Set("備考", "", 1200, "L")
        Call XZMA010_DataGrid_Set("削除", "", 600, "C")
        '
        If GSys.Sys = "FULL" Then
            Call XZMA010_DataGrid_Set("実績共有", "", 1000, "C")
        End If
        If GSys.Lan = True Then
            Call XZMA010_DataGrid_Set("入力中", "入力中", 1200, "L")
            Call XZMA010_DataGrid_Set("処理中", "処理中", 1200, "L")
            Call XZMA010_DataGrid_Set("前回使用", "", 1200, "L")
        End If
        '
        If GSys.Sit = True Then
            Call XZMA010_DataGrid_Set("親会社名", "", 1200, "L")
            Call XZMA010_DataGrid_Set("企業区分", "", 1200, "L")
            
            Call XZMA010_DataGrid_Set("支店コード", "", 1200, "L")
            Call XZMA010_DataGrid_Set("支店名", "", 1200, "L")
        End If
        '
        Call XZMA010_DataGrid_Set("処理開始日付", "処理開始", 2600, "L")
        Call XZMA010_DataGrid_Set("処理終了日付", "処理終了", 2600, "L")
        '
        Call XZMA010_DataGrid_Set("最新処理日", "", 1500, "L")
        Call XZMA010_DataGrid_Set("作成日", "", 1500, "L")
        Call XZMA010_DataGrid_Set("保存日", "", 2600, "L")
        Call XZMA010_DataGrid_Set("復元日", "", 1500, "L")
    Call XZMA010_DataGrid_Action(DataGrid1)
    
    メッセージ = ""
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
AdodcRefresh_ERR:
    pERR_MES = pPROGRAM_ID + "/ AdodcRefresh() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' DataGrid1_Click
'------------------------------------------------
Private Sub DataGrid1_Click()
'
    DoEvents
    
    メッセージ = ""
    Call CEkey.SetFs(企業名Key, True)
End Sub

'------------------------------------------------
' DataGrid1_LostFocus
'------------------------------------------------
Private Sub DataGrid1_LostFocus()
'
    On Error Resume Next
        Dim wCheckValue As Variant
        wCheckValue = Adodc1.Recordset.Fields.Item("企業名Key")
        If Err.Number = 3021 Then GoTo Exit_Sub
    On Error GoTo Err_Hundle
        企業名Key = P8.FCStr(Adodc1.Recordset.Fields.Item("企業名Key"))
    On Error GoTo 0
    
    Call 画面セット(True)
   
    If DataGrid1.Splits.Count <> 1 Then
        DataGrid1.Splits.Remove 1
    End If

Exit_Sub:
    Exit Sub
    '---------------------------------------------------
Err_Hundle:
    If Err.Number = 91 Then Resume Next
    If Err.Number = 94 Then Resume Next
    MsgBox CStr(Err.Number) + ":" + Err.Description
    Resume Exit_Sub
End Sub

'------------------------------------------------
' 画面セット
'------------------------------------------------
Private Function 画面セット(pGridClick As Boolean) As Boolean
'
    Dim wsfmt年月日 As String
'
    On Error GoTo 画面セット_ERR
'
    画面セット = False
'
    ' =========================================
    '                画面クリア
    ' =========================================
    '----------< BUTTON_ENABLE_SET >------------------------------------------------
    運用管理.Caption = "運用管理"
    
    If GProduct <> "金剛石" Then
            B_2.Caption = "": B_2.Enabled = False
    Else
        If GSys.Sit = True Then
            B_2.Caption = "本支店管理": B_2.Enabled = True
        Else
            B_2.Caption = "": B_2.Enabled = False
        End If
    End If
    
    B_3.Caption = "復元処理": B_3.Enabled = True
    B_4.Caption = "DB最適化": B_4.Enabled = True
    B_5.Caption = "特定ﾊﾞｯｸｱｯﾌﾟ": B_5.Enabled = True
    B_6.Caption = "": B_6.Enabled = False
    
    L_最新処理日.Caption = ""
    L_作成日.Caption = ""
    L_保存日時刻.Caption = ""
    L_復元日.Caption = ""
    L_削除日.Caption = ""
    実績 = 0
    削除 = 0
    
    CopyDB名 = wDB名
    wDB名 = ""
    CopyFLG = False
    
    'copy list
    wi決算月 = 3
    wi決算締日 = 31
    wi回収有無 = 0
    wi支払有無 = 0
    wd実績年月 = CDate("2000/01/01")
'
    '----------< DELETE_NEWREC >----------------------------------------------------
    If wNew企業名Key <> "" Then
        Call DELETE_NEWREC
    End If
'
    '----------< FLG 処理 >-----------------------------------------------------
    If w企業名Key <> 企業名Key And wFlg企業名Key = "" And w企業名Key <> "" Then
        Call FLG_NYURYOKUOFF(w企業名Key)
        保存.Caption = "新規企業登録"
        保存.Enabled = False
        pGridClick = False
        w企業名Key = ""
    End If
'
    ' =========================================
    '            企業名マスタ セット
    ' =========================================
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAA070_企業名マスタ"
    wstr = wstr + " Where 企業名Key = '" & 企業名Key & "'"
    Call AdoRecordsetOpen(GDb2, wRs2, wstr)
    If wRs2.eof Then
        wRs2.Close
        Set wRs2 = Nothing
        
        新規変更.Caption = "新規登録"
        保存.Caption = "新規企業登録"
        保存.Enabled = True
        金剛石処理.Enabled = False
        B_4.Enabled = False
        B_5.Enabled = False
        企業名 = ""
        備考.Text = ""
            
        If 企業名Key <> "" Then
            GRet = MsgBox("新規レコードを追加します。よろしいですか？", vbYesNo + vbQuestion)
            If GRet = vbNo Then
                新規変更.Caption = ""
                企業名Key = ""
                Call CEkey.SetFs(企業名Key, True)
            Else
                企業名 = 企業名Key
                If CopyDB名 <> "" Then
                    wstr = "Select *"
                    wstr = wstr + " From DAAA070_企業名マスタ "
                    wstr = wstr + " Where DB名 = '" & CopyDB名 & "'"
                    Call AdoRecordsetOpen(GDb2, wRs2, wstr)
                    If Not wRs2.eof Then
                        If P8.FCStr(wRs2("入力中端末名")) <> "" Or P8.FCStr(wRs2("処理中端末名")) <> "" Then
                            GRet = MsgBox("他端末で変更前データを使用中の為" + vbCrLf + _
                                "変更前データを基準にデータを作成できません", vbOKOnly + vbExclamation)
                        
                            新規変更.Caption = ""
                            企業名Key = ""
                            企業名 = ""
                            Call CEkey.SetFs(企業名Key, True)
                        
                            wRs2.Close
                            Set wRs2 = Nothing
                            
                            Exit Function
                        Else
                            GRet = MsgBox("変更前データを基準にデータを作成しますか？", vbYesNo + vbQuestion)
                            If GRet = vbYes Then
                                CopyFLG = True
                        
                                If GSys.Sys = "FULL" Then
                                    実績 = wRs2("実績データ共有")
                                End If
                            Else
                                実績 = 0
                            End If
                            
                            'copy
                            wi決算月 = wRs2("決算月")
                            wi決算締日 = wRs2("決算締日")
                            wi回収有無 = wRs2("回収有無")
                            wi支払有無 = wRs2("支払有無")
                            wd実績年月 = Format(wRs2("最終実績年月"), "yyyy/mm/dd")
                            
                        End If
                    End If
                    wRs2.Close
                    Set wRs2 = Nothing
                End If
                
                '----------< FLG処理 >----------------------------------------------
                GRet = FLG_TOROKUSET(企業名Key, "新規登録")
                If GRet <> True Then
                    Call CEkey.SetFs(企業名Key, True)
                    Exit Function
                End If
                
                Call CEkey.SetFs(企業名, True)
            End If
        Else
            保存.Enabled = False
        End If
    Else
        画面セット = True
        
        新規変更.Caption = "選択／内容変更"
        If 保存.Caption = "変更内容登録" Then
            保存.Enabled = True
        End If
        
        If (P8.FCStr(wRs2("入力中端末名")) = "" Or wRs2("入力中端末名") = GMyComputerName) _
            And (P8.FCStr(wRs2("処理中端末名")) = "" Or wRs2("処理中端末名") <> GMyComputerName) Then
            金剛石処理.Enabled = True
            B_4.Enabled = True
            B_5.Enabled = True
                
            wFlg企業名Key = ""
        Else
            金剛石処理.Enabled = False
            B_4.Enabled = False
            B_5.Enabled = False
        
            wFlg企業名Key = 企業名Key
        End If
        
        企業名 = P8.FCStr(wRs2("企業名"))
        備考.Text = P8.FCStr(wRs2("備考"))
        '
        If GSys.Sys = "FULL" Then
            実績 = wRs2("実績データ共有")
        End If
        '
        削除 = IIf(P8.FCStr(wRs2("削除日")) = "", 0, 1)
'
        If Gfmt年月日 = "" Then
            If G基本情報.日付入力区分 = "0" Then
                '和暦入力
                wsfmt年月日 = "ee年mm月dd日"
            Else
                wsfmt年月日 = "yyyy/mm/dd"
                '西暦入力
            End If
        Else
            wsfmt年月日 = Gfmt年月日
        End If
'
        L_最新処理日.Caption = Format(P8.FCStr(wRs2("最新処理日")), wsfmt年月日 & " hh:nn:ss")
        L_作成日.Caption = Format(P8.FCStr(wRs2("作成日")), wsfmt年月日 & " hh:nn:ss")
        L_保存日時刻.Caption = Format(P8.FCStr(wRs2("保存日")), wsfmt年月日 & " hh:nn:ss")
        L_復元日.Caption = Format(P8.FCStr(wRs2("復元日")), wsfmt年月日 & " hh:nn:ss")
        L_削除日.Caption = Format(P8.FCStr(wRs2("削除日")), wsfmt年月日 & " hh:nn:ss")
    
        wDB名 = P8.FCStr(wRs2("DB名"))
        w企業名Key = 企業名Key
        
        wRs2.Close
        Set wRs2 = Nothing
    End If
    
    '------------------------------------------
    '          ** グリッドコントロール **
    '------------------------------------------
    If Not pGridClick Then
        DoEvents
        Call AdodcRefresh
    End If
    
    DoEvents
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "企業名Key = '" + 企業名Key + "'")
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
画面セット_ERR:
    pERR_MES = pPROGRAM_ID + "/ 画面セット() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

'------------------------------------------------
' 企業名Key_GotFocus
'------------------------------------------------
Private Sub 企業名Key_GotFocus()
    Call CEkey.AllSelect
End Sub

'------------------------------------------------
' 企業名Key_LostFocus
'------------------------------------------------
Private Sub 企業名Key_LostFocus()
'
    Dim ws01 As String
'
    On Error GoTo 企業名Key_LostFocus_ERR
'
    Call P8.FCControlLeft(企業名Key, 50)
    
    Select Case Screen.ActiveControl.Name
    Case "DataGrid1", "企業名Key", "検索備考", "検索", "最新GRID", "閉じる", "運用管理", _
         "B_2", "B_3", "B_4", "B_5", "B_6"
        Exit Sub
    End Select
   
    If 企業名Key = "" And Screen.ActiveForm.Name = "FAA020_ログインユーザ選択" Then
        MsgBox "コードを入力してください"
        Call CEkey.SetFs(企業名Key, True)
        
        Exit Sub
    End If
    
    ws01 = StrConv(企業名Key, VbStrConv.vbNarrow)
    ws01 = LCase(ws01)
    If ws01 = "backup" & 企業名Key Then
        MsgBox "企業名Keyにbackupは使用しないでください"
        Call CEkey.SetFs(企業名Key, True)
        
        Exit Sub
    ElseIf ws01 = "k000" Or ws01 = "list" Or ws01 = "金剛石" _
            Or ws01 = "金剛石wk" Or ws01 = "金剛石変更" Or ws01 = "listwk" Or ws01 = "list変更" Then
        MsgBox "企業名Keyにシステム用mdbは使用しないでください"
        Call CEkey.SetFs(企業名Key, True)
        
        Exit Sub
    End If
    
    GRet = 禁則文字(企業名Key)
    If GRet <> True Then
        MsgBox "企業名Keyに禁則文字 (\ / : , ; * ? "" < > |) は使用できません"

        Call CEkey.SetFs(企業名Key, True)
        Exit Sub
    End If
'
    '----------< FLG 処理 >-----------------------------------------------------
    GRet = FLG_TOROKUSET(企業名Key)
    If GRet = True Then
        FLG_LOSTK = True
        保存.Caption = "変更内容登録"
    Else
        Exit Sub
'        '----------< RESET Adodc >--------------------------------------------------
'        Call ADODC_RESET
    End If
    
    DoEvents
    
    Call 画面セット(False)
    Call CEkey.AllSelect
    
    DoEvents
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
企業名Key_LostFocus_ERR:
    pERR_MES = pPROGRAM_ID + "/ 企業名Key_LostFocus() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' 登録後初期セット
'------------------------------------------------
Private Sub 登録後初期セット()
'
    Dim w検索備考 As String
'
    FLG_LOSTK = False
    
    保存.Caption = "新規企業登録"
'    wDB名 = ""
    w企業名Key = ""
    wFlg企業名Key = ""
    wNew企業名Key = ""
    企業名Key = ""
    
    Call 画面セット(False)
    新規変更.Caption = ""
    
    '----------------------------------------
    '               更新行を表示
    '----------------------------------------
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "企業名Key = '" + 企業名Key + "'")
    Call CEkey.SetFs(企業名Key, True)

    '----------------------------------------
    '            備考セット
    '----------------------------------------
    w検索備考 = 検索備考
    検索備考 = ""
    
    備考.Clear
    検索備考.Clear
    
    備考.AddItem ""
    検索備考.AddItem ""
    
    wstr = ""
    wstr = wstr + "Select 備考"
    wstr = wstr + " From DAAA070_企業名マスタ"
    wstr = wstr + " Where 備考 <> '' "
    wstr = wstr + " Group By 備考"
    wstr = wstr + " Order By 備考"
    Call AdoRecordsetOpen(GDb2, wRs2, wstr)
        Do Until wRs2.eof
            備考.AddItem (P8.FCStr(wRs2("備考")))
            検索備考.AddItem (P8.FCStr(wRs2("備考")))
                         
            wRs2.MoveNext
        Loop
    wRs2.Close
    Set wRs2 = Nothing

    検索備考.Text = w検索備考
'
End Sub

'------------------------------------------------
' ADODC_RESET
'------------------------------------------------
Private Sub ADODC_RESET()
'
    On Error GoTo ADODC_RESET_ERR
'
    DoEvents
'
    '----------< DataGrid Close >----------------------------------------------
    Set DataGrid1.DataSource = Nothing
    Adodc1.Recordset.Close
'
    '----------< List.mdb Close >-----------------------------------------------
    GDb2.Close
    Set GDb2 = Nothing
'
    Call BUTTON_ENABLE_SET(True)
'
    検索備考.Text = ""
    FLG_KIGYO = False
    DoEvents
'
    Call Form_Load
'
    DoEvents
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
ADODC_RESET_ERR:
    pERR_MES = pPROGRAM_ID + "/ ADODC_RESET() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' BUTTON_ENABLE_SET
'------------------------------------------------
Private Sub BUTTON_ENABLE_SET(pEnabled As Boolean)
'
    If pEnabled = False Then
        DataGrid1.Enabled = False
        保存.Enabled = False
        金剛石処理.Enabled = False
        最新GRID.Enabled = False
        運用管理.Enabled = False
        B_2.Enabled = False
        B_3.Enabled = False
        B_4.Enabled = False
        B_5.Enabled = False
        B_6.Enabled = False
        閉じる.Enabled = False
    Else
        DataGrid1.Enabled = True
'        保存.Enabled = True
'        金剛石処理.Enabled = True
        最新GRID.Enabled = True
        運用管理.Enabled = True
        B_2.Enabled = True
        '@001 ADD
        'B_3.Enabled = True
'        B_4.Enabled = True
'        B_5.Enabled = True
        B_6.Enabled = True
        閉じる.Enabled = True
    End If
    
    DoEvents
'
End Sub

'------------------------------------------------
' BUTTON_TOOLTIPTEXT_SET
'------------------------------------------------
Private Sub BUTTON_TOOLTIPTEXT_SET()
'
    '----------< ToolTrip Text >----------------------------------------------------
    DataGrid1.ToolTipText = "登録済データをマウスで選択します。"
    
    金剛石処理.ToolTipText = "選択企業名Keyの" & GProduct & "の処理開始メニューへ。"
    閉じる.ToolTipText = GProduct & "処理を終了します。"

    '----------< ToolTipText >------------------------------------------------------
    If 運用管理.Caption = "運用管理" Then
        運用管理.ToolTipText = "運用管理メニューへ。"
    ElseIf 運用管理.Caption = "前メニュー" Then
        運用管理.ToolTipText = "前メニューへ。"
    Else
        運用管理.ToolTipText = ""
    End If
    
    If B_2.Caption = "本支店管理" Then
        B_2.ToolTipText = "本支店管理メニューへ。"
    ElseIf B_2.Caption = "フラグ解除" Then
        B_2.ToolTipText = "LAN対応時のリセット処理。"
    Else
        B_2.ToolTipText = ""
    End If
    
    If B_3.Caption = "復元処理" Then
        B_3.ToolTipText = "企業名Key単位での復元処理。"
    ElseIf B_3.Caption = "CSV出力先指定" Then
        B_3.ToolTipText = "CSV出力先フォルダー指定"
    Else
        B_3.ToolTipText = ""
    End If
    
    If B_4.Caption = "DB最適化" Then
        B_4.ToolTipText = "データベースの最適化処理。企業名Keyを選択後クリックしてください。"
    ElseIf B_4.Caption = "サーバー指定" Then
        B_4.ToolTipText = "LAN対応時のデータサーバー指定。"
    Else
        B_4.ToolTipText = ""
    End If
    
    If B_5.Caption = "特定ﾊﾞｯｸｱｯﾌﾟ" Then
        B_5.ToolTipText = "企業名Key単位でのバックアップ処理。企業名Keyを選択後クリックしてください。"
    ElseIf B_5.Caption = "全体ﾊﾞｯｸｱｯﾌﾟ" Then
        B_5.ToolTipText = "フォルダーのバックアップ処理。"
    Else
        B_5.ToolTipText = ""
    End If
    
    If B_6.Caption = "ログ一覧表" Then
        B_6.ToolTipText = GProduct & "使用ログ。"
    ElseIf B_6.Caption = "完全削除" Then
        B_6.ToolTipText = "削除マークが付いている企業名KeyのDBを削除します。"
    Else
        B_6.ToolTipText = ""
    End If
'
    DoEvents
'
End Sub

'------------------------------------------------
' LostFocus
'------------------------------------------------
Private Sub 検索備考_LostFocus()
    Call P8.FCControlLeft(検索備考, 50)
End Sub

Private Sub 備考_LostFocus()
    Call P8.FCControlLeft(備考, 50)
End Sub
'
'##################################################################################
'#
'#                                ＡｆｔｅｒＣｌｉｃｋ
'#
'##################################################################################
'------------------------------------------------
' 保存_Click
'------------------------------------------------
Private Sub 保存_Click()
'
    Dim wDb As New ADODB.Connection
    Dim wJET As New JetEngine
    
    Dim wSDate As String, wsDB名 As String
    Dim ws01 As String, ws02 As String
    
    Dim j As Integer
'
    On Error GoTo 保存_Click_ERR
'
    wSDate = Format(Now, "yyyy/mm/dd hh:mm:ss")
    
    If P8.FCStr(企業名.Text) = "" Then
        MsgBox "企業名を入力してください"
        
        Call CEkey.SetFs(企業名, True)
        Exit Sub
    ElseIf False = 禁則文字(企業名.Text) Then
        MsgBox "企業名に禁則文字 (\ / : , ; * ? "" < > |) は使用できません"
        
        Call CEkey.SetFs(企業名, True)
        Exit Sub
    End If
'
    If wDB名 = "" Then
        wsDB名 = P8.FCStr(企業名Key.Text) & ".mdb"
    Else
        wsDB名 = wDB名
    End If
'
    '----------< List.mdb >---------------------------------------------------------
    '----------< DAAA070_企業名マスタ >---------------------------------------------
    wstr = "Select *"
    wstr = wstr + " From DAAA070_企業名マスタ"
    wstr = wstr + " Where 企業名Key = '" & 企業名Key.Text & "'"
    wstr = wstr + " And DB名 <> '" & wsDB名 & "'"
    Call AdoRecordsetOpen(GDb2, wRs2, wstr)
    If Not wRs2.eof Then
        MsgBox "企業名Keyが重複しています"
        wRs2.Close
        Set wRs2 = Nothing
        
        '----------< FLG 処理 >-----------------------------------------------------
        Call FLG_NYURYOKUOFF(企業名Key)
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        
        Exit Sub
    End If
    wRs2.Close
    Set wRs2 = Nothing
'
    '----------< List.mdb >---------------------------------------------------------
    '----------< DAAA070_企業名マスタ >---------------------------------------------
    GRet = LISTDATA_UPDATE(wSDate)
    If GRet <> True Then
        Exit Sub
    End If
'
    '----------< BUTTON_ENABLE_SET >------------------------------------------------
    Call BUTTON_ENABLE_SET(False)
'
    ' =========================================
    '       ﾃﾝﾌﾟﾚｰﾄよりﾌｧｲﾙの複製と最適化
    ' =========================================
    If 保存.Caption = "新規企業登録" Then
        If CopyFLG = True Then
            GRet = FileMaker(GSerDir & "\" & CopyDB名, GSerDir & "\" & wDB名)
        Else
            GRet = FileMaker(GSerDir & "\" & GTemp, GSerDir & "\" & wDB名)
        End If
        If GRet <> True Then
            GRet = MsgBox("他端末で使用中です" + vbCrLf + "使用可能になった時点で再度実行してください", vbExclamation + vbOKOnly)
            
            '----------< Delete List data >-----------------------------------------
            wstr = ""
            wstr = wstr + "Delete"
            wstr = wstr + " From DAAA070_企業名マスタ"
            wstr = wstr + " Where DB名 = '" + wDB名 + "'"
            GDb2.Execute wstr
            DoEvents
            
            ' =========================================
            '                   後処理
            ' =========================================
            Call ADODC_RESET
'            Call BUTTON_ENABLE_SET(True)
            MsgBox "新規企業登録できませんでした", vbExclamation
            
            Exit Sub
        End If
    End If
    
    ' =========================================
    '           テーブル 更新
    ' =========================================
    wstr = ""
    '----------< KXXX.mdb Open >----------------------------------------------------
    Call AdoDbOpen("Jet", wDb, GSerDir + "\" + wDB名, "", , GPwd)
        If FLG_New = True Then
            '----------< DAAA070_企業名マスタ >--------------------------------------
            '----------< Clear DAAA070_企業名マスタ >-------------------------------
            wstr = "Delete * From DAAA070_企業名マスタ"
            wDb.Execute wstr
            
            wstr = "INSERT INTO DAAA070_企業名マスタ"
            wstr = wstr + " (企業名Key,企業名,備考,最新処理日,作成日,DB名,削除日,実績データ共有)"
            wstr = wstr + " Values ("
            wstr = wstr + "'" + 企業名Key.Text + "',"
            wstr = wstr + "'" + 企業名.Text + "',"
            wstr = wstr + "'" + 備考.Text + "',"
            wstr = wstr + "#" + wSDate + "#,"
            wstr = wstr + "#" + wSDate + "#,"
            wstr = wstr + "'" + wDB名 + "',"
            
            Select Case 削除.Value
            Case 0
                wstr = wstr + "Null,"
            Case 1
                wstr = wstr + "#" + wSDate + "#,"
            End Select
            
            If GSys.Sys = "LUFU" Then
                wstr = wstr + "0);"
            Else
                Select Case 実績.Value
                Case 0
                    wstr = wstr + "0);"
                Case 1
                    wstr = wstr + "1);"
                End Select
            End If
            
            wDb.Execute wstr
                        
            '----------< DELETE >---------------------------------------------------
            '----------< DBBA010_売上実績 >-----------------------------------------
            If 実績.Value = 0 Or GSys.Sys = "LUFU" Then
                If CopyDB名 = GTemp Then
                    wstr = "Delete * From DBBA010_売上実績"
                    wDb.Execute wstr
                End If
            End If
            
            '----------< UPDATE >---------------------------------------------------
            '----------< DBBA010_売上実績 >-----------------------------------------
            wstr = "UPDATE DBBA010_売上実績"
            wstr = wstr + " SET"
            wstr = wstr + " 販売データ有無=0"
            wDb.Execute wstr
            
            '----------< DBB010_受注発注 >-----------------------------------------
            wstr = "Delete * From DBB010_受注発注"
            wDb.Execute wstr
        
            '----------< DBB010_売上実績販売 >--------------------------------------
            wstr = "Delete * From DBB010_売上実績販売"
            wDb.Execute wstr
        
            '----------< DBBA010_本部経費振替 >-------------------------------------
            wstr = "Delete * From DBBA010_本部経費振替"
            wDb.Execute wstr
        
            '----------< DBBA010_基幹データ調整 >-------------------------------------
            wstr = "Delete * From DBBA010_基幹データ調整"
            wDb.Execute wstr
        
        
            '----------< UPDATE >---------------------------------------------------
            '----------< DAAA010_基本情報 >-----------------------------------------
            wstr = "UPDATE DAAA010_基本情報"
            wstr = wstr + " SET"
            wstr = wstr + " 支店コード = '',"
            wstr = wstr + " 支店名 = '単独企業',"
            wstr = wstr + " 企業区分 = '単独企業',"
            wstr = wstr + " 資金調達区分 = '2',"
            wstr = wstr + " 有担銀行 = 'ZZ',"
            wstr = wstr + " 無担銀行 = 'ZZ'"
            wstr = wstr + " WHERE System='System'"
            wDb.Execute wstr
            
        Else
            '----------< DAAA070_企業名マスタ >--------------------------------------
            wstr = "Update "
            wstr = wstr + "  DAAA070_企業名マスタ "
            wstr = wstr + "Set "
            wstr = wstr + "企業名 = '" + 企業名.Text + "',"
            wstr = wstr + "備考 = '" + 備考.Text + "',"
            wstr = wstr + "最新処理日 = #" + wSDate + "#,"
                
            Select Case 削除.Value
            Case 0
                wstr = wstr + "削除日 = Null,"
            Case 1
                wstr = wstr + "削除日 = #" + wSDate + "#,"
            End Select
            
            If GSys.Sys = "LUFU" Then
                wstr = wstr + "実績データ共有 =0"
            Else
                Select Case 実績.Value
                Case 0
                    wstr = wstr + "実績データ共有 =0"
                Case 1
                    wstr = wstr + "実績データ共有 =1"
                End Select
            End If

            wstr = wstr + " Where 企業名Key = '" & 企業名Key.Text & "'"
            wDb.Execute wstr
        End If
'
    '----------< KXXX.mdb Close >---------------------------------------------------
    wDb.Close
    Set wDb = Nothing
'
    ' =========================================
    '           KIGYOSHORI_END
    ' =========================================
    Call KIGYOSHORI_END(企業名.Text, wSDate)
    
    '----------< KXXX.mdb Open >----------------------------------------------------
    Call AdoDbOpen("Jet", wDb, GSerDir + "\" + wDB名, "", , GPwd)

    '----------< KXXX.mdb >---------------------------------------------------------
    '----------< DAAA070_企業名マスタ >---------------------------------------------
    wstr = "Update"
    wstr = wstr + " DAAA070_企業名マスタ"
    wstr = wstr + " Set"
    wstr = wstr + " 処理終了日付 = #" + wSDate + "#,"
    wstr = wstr + " 保存日 = #" + wSDate + "#,"
    wstr = wstr + " 入力中端末名 = '',"
    wstr = wstr + " 処理中端末名 = ''"
    wstr = wstr + " Where 企業名Key = '" & 企業名.Text & "'"
    wDb.Execute wstr
    
    '----------< KXXX.mdb Close >---------------------------------------------------
    wDb.Close
    Set wDb = Nothing
'
    '----------< LOG WRITE >--------------------------------------------------------
    If 保存.Caption = "新規企業登録" Then
        GStr = "1," & "1," & P8.FCChr(GMyComputerName, 16) & "," & P8.FCChr(企業名Key, 30) & "," & P8.FCChr(企業名, 30)
    Else
        GStr = "1," & "2," & P8.FCChr(GMyComputerName, 16) & "," & P8.FCChr(企業名Key, 30) & "," & P8.FCChr(企業名, 30)
    End If
    GRet = PUT_LOG_FILE(GStr)
'
    ' =========================================
    '                   後処理
    ' =========================================
    Call ADODC_RESET
'    Call BUTTON_ENABLE_SET(True)
    MsgBox "企業情報の登録・更新処理が完了しました", vbInformation
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
保存_Click_ERR:
    pERR_MES = pPROGRAM_ID + "/ 保存_Click() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' LISTDATA_UPDATE
'------------------------------------------------
Private Function LISTDATA_UPDATE(pSdate As String) As Boolean
'
    Dim wi01 As Integer
    Dim ws01 As String, ws02 As String
    Dim wsDB名 As String
'
    On Error GoTo LISTDATA_UPDATE_ERR
'
    LISTDATA_UPDATE = False
'
LISTDATA_UPDATE_ERR_RETRY:
    '----------< DAAA070_企業名マスタ >---------------------------------------------
    wsDB名 = 企業名Key.Text & ".mdb"
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAA070_企業名マスタ"
    wstr = wstr + " Where 企業名Key = '" & 企業名Key.Text & "'"
    Call AdoRecordsetOpen(GDb2, wRs2, wstr)
    If Not wRs2.eof Then
        If wRs2("企業名") = 企業名.Text _
            And wRs2("備考") = 備考.Text _
            And wRs2("実績データ共有") = 実績 _
            And ((P8.FCStr(wRs2("削除日")) = "" And 削除.Value = 0) _
             Or (P8.FCStr(wRs2("削除日")) <> "" And 削除.Value = 1)) Then
             
            MsgBox "内容が変更されていません"
            wRs2.Close
            Set wRs2 = Nothing
            
            '----------< FLG 処理 >-------------------------------------------------
            Call FLG_NYURYOKUOFF(企業名Key)
            '----------< RESET Adodc >----------------------------------------------
            Call ADODC_RESET
        
            Exit Function
        End If
        
        If Dir(GSerDir + "\" + wDB名) = "" Then
            wRs2.Close
            Set wRs2 = Nothing
                
            MsgBox "対象MDBが見つかりませんエクスプローラー側から削除、または移動された可能性があります｡ ", vbCritical
            GRet = MsgBox("対象をリストから抹消しますか？", vbYesNo + vbQuestion)
            If GRet = vbYes Then
                '----------< Delete List data >-------------------------------------
                wstr = ""
                wstr = wstr + "Delete "
                wstr = wstr + "From DAAA070_企業名マスタ "
                wstr = wstr + "Where DB名 = '" + wDB名 + "'"
                GDb2.Execute (wstr)
            Else
                '----------< FLG 処理 >---------------------------------------------
                Call FLG_NYURYOKUOFF(企業名Key)
            End If
            
            Exit Function
        End If
        
        '----------< FLG 処理 >-----------------------------------------------------
        '----------< FLG_KEYUPDATE >------------------------------------------------
        wi_Kosin2 = P8.FCDbl(wRs2("更新回数"))
        If wi_Kosin1 <> wi_Kosin2 Then
            '----------< Msg >------------------------------------------------------
            ws01 = "他端末で更新されました"
            ws02 = "この端末(" & GMyComputerName & ")では更新できません"
            GRet = MsgBox(ws01 & vbCrLf & ws02, vbOKOnly + vbExclamation)
            wRs2.Close
            Set wRs2 = Nothing
                
            Exit Function
        End If
        
        ws01 = P8.FCStr(wRs2("入力中端末名"))
        If ws01 <> "" And ws01 <> GMyComputerName Then
            '----------< Msg >------------------------------------------------------
            If ws01 = "" Then
                ws02 = "他端末で同一企業(" & 企業名Key.Text & ")入力中です"
            Else
                ws02 = "他端末(" & ws01 & ")で同一企業(" & 企業名Key.Text & ")入力中です"
            End If
            GRet = MsgBox(ws02 & vbCrLf & Msg_01, vbOKOnly + vbExclamation)
            wRs2.Close
            Set wRs2 = Nothing
                
            Exit Function
        End If
    
        ws01 = P8.FCStr(wRs2("処理中端末名"))
        If ws01 <> "" And ws01 <> GMyComputerName Then
            wRs2.Close
            Set wRs2 = Nothing
            
            '----------< Msg >-------------------------------------------------------
            If ws01 = "" Then
                ws02 = "他端末で同一企業(" & 企業名Key.Text & ")入力中です"
            Else
                ws02 = "他端末(" & ws01 & ")で同一企業(" & 企業名Key.Text & ")入力中です"
            End If
            GRet = MsgBox(ws02 & vbCrLf & Msg_01, vbOKOnly + vbExclamation)
                
            Exit Function
        End If
    
        wRs2("企業名") = P8.FCStr(企業名.Text)
        wRs2("備考") = P8.FCStr(備考.Text)
        
        If GSys.Sys = "FULL" Then
            wRs2("実績データ共有") = 実績
        Else
            wRs2("実績データ共有") = 0
        End If
        
        'copy
        If FLG_New = True Then
            wRs2("決算月") = wi決算月
            wRs2("決算締日") = wi決算締日
            wRs2("回収有無") = wi回収有無
            wRs2("支払有無") = wi支払有無
            wRs2("最終実績年月") = Format(wd実績年月, "yyyy/mm/dd")
        End If

        Select Case 削除.Value
        Case 0
            wRs2("削除日") = Null
        Case 1
            If L_削除日.Caption = "" Then
                wRs2("削除日") = P8.FCDate(pSdate)
            End If
        End Select
        
        wRs2("最新処理日") = P8.FCDate(pSdate)
        wDB名 = P8.FCStr(wRs2("DB名"))
        '----------< FLG処理 >------------------------------------------------------
        wRs2("更新回数") = wi_Kosin2 + 1
        '
        wRs2.Update
    End If
    wRs2.Close
    Set wRs2 = Nothing
    
    wNew企業名Key = ""

On Error GoTo 0
'
    LISTDATA_UPDATE = True
'
    Exit Function
'
'----------< ERROR ROUTINE >--------------------------------------------------------
LISTDATA_UPDATE_ERR:
    If Err.Number = -214727887 Or Err.Number = -2147217887 Then
        Sleep (1000)
        Resume LISTDATA_UPDATE_ERR_RETRY
    End If
    '
    pERR_MES = pPROGRAM_ID + "/ LISTDATA_UPDATE() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                "他端末で同一企業(" & 企業名Key.Text & ")使用中です"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)
    '
    '----------< ADODC_RESET >------------------------------------------------------
    Call ADODC_RESET
    '
    Resume LISTDATA_UPDATE_ERR_END
LISTDATA_UPDATE_ERR_END:
    Exit Function
'
End Function

'------------------------------------------------
' 金剛石処理_Click
'------------------------------------------------
Private Sub 金剛石処理_Click()
'
    Dim wDb As New ADODB.Connection
    Dim wSDate As String, wsDB名 As String
    Dim ws01 As String
'
    On Error GoTo 金剛石処理_Click_ERR
'
    If 企業名Key.Text = "" And 企業名.Text = "" And 備考.Text = "" And L_最新処理日.Caption = "" _
      And L_保存日時刻.Caption = "" And L_復元日.Caption = "" And L_削除日.Caption = "" Then
        MsgBox "企業が選択されていません"
        Call CEkey.SetFs(企業名Key, True)
        Exit Sub
    End If
'
    '----------< BUTTON_ENABLE_SET >------------------------------------------------
    Call BUTTON_ENABLE_SET(False)
'
    wstr = ""
    If Dir(GSerDir + "\" + wDB名) = "" Then
        MsgBox "対象MDBが見つかりません。エクスプローラー側から削除、または移動された可能性があります｡ ", vbCritical
        GRet = MsgBox("対象をリストから抹消しますか？", vbYesNo + vbExclamation)
        If GRet = vbYes Then
            '----------< BUTTON_ENABLE_SET >----------------------------------------
            Call BUTTON_ENABLE_SET(False)
            
            '----------< Delete List data >-----------------------------------------
            wstr = ""
            wstr = wstr + "Delete "
            wstr = wstr + "From DAAA070_企業名マスタ "
            wstr = wstr + "Where"
            wstr = wstr + "   DB名 = '" + wDB名 + "'"
            GDb2.Execute (wstr)
            
            ' =========================================
            '                   後処理
            ' =========================================
            Call ADODC_RESET
'            Call BUTTON_ENABLE_SET(True)
        End If
        Exit Sub
    End If
'
    wSDate = Format(Now, "yyyy/mm/dd hh:mm:ss")
    '----------< FLG 処理 >---------------------------------------------------------
    GRet = FLG_KIGYOSHORI(企業名Key, wSDate)
    If GRet <> True Then
        Exit Sub
    End If
'
    GDbName = GSerDir + "\" + wDB名
'
    '----------< AdoDbOpen_Check >--------------------------------------------------
    GRet = ADODBOPEN_CHECK("Jet", GDb, GDbName, "", , GPwd, "排他")
    If GRet <> True Then
        GRet = MsgBox("他端末で使用中です" + vbCrLf + "使用可能になった時点で再度実行してください", vbExclamation + vbOKOnly)
        GDbName = ""
        
        '----------< FLG 処理 >-----------------------------------------------------
        Call FLG_NYURYOKUOFF(企業名Key)
        '
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        '
        Exit Sub
    End If
'
    '----------< KXXX Ver Check >---------------------------------------------------
    '----------< KXXX.mdb Open >----------------------------------------------------
    Call AdoDbOpen("Jet", wDb, GDbName, "", , GPwd)
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAA000_バージョン"
    wstr = wstr + " Where System = 'System'"
    Call AdoRecordsetOpen(wDb, wRs2, wstr)
    If Not wRs2.eof Then
        ws01 = P8.FCStr(wRs2("Version"))
    End If
    wRs2.Close
    Set wRs2 = Nothing
    '----------< KXXX.mdb Close >---------------------------------------------------
    wDb.Close
    Set wDb = Nothing
    
    If GVerNo <> ws01 Then
        GRet = MsgBox("バージョンの違うMDBを参照しています", vbExclamation + vbOKOnly)
        GDbName = ""
        
        '----------< FLG 処理 >-----------------------------------------------------
        Call FLG_NYURYOKUOFF(企業名Key)
        '
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        '
        Exit Sub
    End If
'
    '----------< KIGYOSHORI_STR >---------------------------------------------------
    '----------< List.mdb Open >----------------------------------------------------
    ws01 = "": G実績共有 = ""
    wsDB名 = ""
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAA070_企業名マスタ"
    wstr = wstr + " Where 企業名Key = '" & 企業名Key & "'"
    Call AdoRecordsetOpen(GDb2, wRs2, wstr)
    If Not wRs2.eof Then
        wsDB名 = wRs2("DB名")
        If GSys.Sys = "FULL" Then
            ws01 = P8.FCStr(wRs2("実績データ共有"))
        End If
    
        wRs2("入力中端末名") = ""
        
        wRs2.Update
    End If
    wRs2.Close
    Set wRs2 = Nothing
    
    If ws01 = "1" Then
        G実績共有 = "共有"
    Else
        G実績共有 = "単独"
    End If

    '----------< KXXX.mdb Open >----------------------------------------------------
    Call AdoDbOpen("Jet", wDb, GSerDir + "\" + wsDB名, "", , GPwd)
        wstr = "Update "
        wstr = wstr + "DAAA070_企業名マスタ"
        wstr = wstr + " Set"
        wstr = wstr + " 最新処理日 = #" + wSDate + "#,"
        wstr = wstr + " 処理開始日付 = #" + wSDate + "#,"
        wstr = wstr + " 処理終了日付 = #" + wSDate + "#,"
        wstr = wstr + " 端末コンピュータ名 = '" + GMyComputerName + "'"
        wstr = wstr + " Where 企業名Key = '" & 企業名Key.Text & "'"
        wDb.Execute wstr
    '----------< KXXX.mdb Close >---------------------------------------------------
    wDb.Close
    Set wDb = Nothing
'
    '----------< LOG WRITE >--------------------------------------------------------
    GStr = "2," & "9," & P8.FCChr(GMyComputerName, 16) & "," & P8.FCChr(企業名Key, 30) & "," & P8.FCChr(企業名, 30)
    GRet = PUT_LOG_FILE(GStr)
'
    DataGrid1.Visible = False
    Set DataGrid1.DataSource = Nothing
    Adodc1.Recordset.Close
    
    '----------< List.mdb Close >---------------------------------------------------
    GDb2.Close
    Set GDb2 = Nothing
'
    FLG_KIGYO = True
'
    DoEvents
'
    GKeyName = P8.FCStr(企業名Key.Text)
    
'
    frm_Fログイン.Show
'    FAA015_メインメニュー借換たろう.Show vbModal
    
    Unload Me
'    '金剛石 or 借換たろう！
'    If GProduct <> "金剛石" Then
'        FAA015_メインメニュー借換たろう.Show vbModal
'    Else
'        'FAA010_メインメニュー.Show vbModal
'    End If
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
金剛石処理_Click_ERR:
    pERR_MES = pPROGRAM_ID + "/ 金剛石処理_Click() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' 検索_Click
'------------------------------------------------
Private Sub 検索_Click()
    Call 登録後初期セット
End Sub

'------------------------------------------------
' 最新GRID_Click
'------------------------------------------------
Private Sub 最新GRID_Click()
'
    '----------< DELETE_NEWREC >----------------------------------------------------
    If wNew企業名Key <> "" Then
        Call DELETE_NEWREC
    End If
'
    '----------< FLG 処理 >---------------------------------------------------------
    GRet = FLG_CHECK_NYURYOKUOFF(企業名Key)
'
    '----------< RESET Adodc >------------------------------------------------------
    Call ADODC_RESET
End Sub

'------------------------------------------------
' 運用管理_Click
'------------------------------------------------
Private Sub 運用管理_Click()
'
    メッセージ = ""
    
    If 運用管理.Caption = "前メニュー" Then
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        
        Call BUTTON_TOOLTIPTEXT_SET
        
        Exit Sub
    End If
'
    '----------< DELETE_NEWREC >----------------------------------------------------
    If wNew企業名Key <> "" Then
        Call DELETE_NEWREC
    End If
'
    '----------< FLG 処理 >---------------------------------------------------------
    GRet = FLG_CHECK_NYURYOKUOFF(企業名Key)
'
    If 運用管理.Caption = "運用管理" Then
        運用管理.Caption = "前メニュー"
        B_2.Caption = "フラグ解除": B_2.Enabled = True
        
        '@001 ADD
        'B_3.Caption = "": B_3.Enabled = False
        If GSys.Lan = True Then
            B_3.Caption = "CSV出力先指定": B_4.Enabled = True
        Else
            B_3.Caption = "": B_3.Enabled = False
        End If
        
        If GSys.Lan = True Then
            B_4.Caption = "サーバー指定": B_4.Enabled = True
        Else
            B_4.Caption = "": B_4.Enabled = False
        End If
        
        B_5.Caption = "全体ﾊﾞｯｸｱｯﾌﾟ": B_5.Enabled = True
        B_6.Caption = "完全削除": B_6.Enabled = True
            
        Call BUTTON_TOOLTIPTEXT_SET
    End If
'
End Sub

'------------------------------------------------
' B_2_Click
'------------------------------------------------
Private Sub B_2_Click()
    メッセージ = ""
    
    If B_2.Caption = "本支店管理" Then
        Call LISTDB処理
    ElseIf B_2.Caption = "フラグ解除" Then
        Call フラグ解除
    End If
End Sub

'------------------------------------------------
' B_3_Click
'------------------------------------------------
Private Sub B_3_Click()
    メッセージ = ""
    
    If B_3.Caption = "復元処理" Then
        Call 復元処理
    '@001 ADD
    ElseIf B_3.Caption = "CSV出力先指定" Then
        Call CSV出力先指定
    End If
End Sub

'------------------------------------------------
' B_4_Click
'------------------------------------------------
Private Sub B_4_Click()
    メッセージ = ""
    
    If B_4.Caption = "DB最適化" Then
        Call 最適化
    ElseIf B_4.Caption = "サーバー指定" Then
        Call サーバー指定
    End If
End Sub

'------------------------------------------------
' B_5_Click
'------------------------------------------------
Private Sub B_5_Click()
    メッセージ = ""
    
    If B_5.Caption = "特定ﾊﾞｯｸｱｯﾌﾟ" Then
        Call バックアップ
    ElseIf B_5.Caption = "全体ﾊﾞｯｸｱｯﾌﾟ" Then
        Call システムバックアップ
    End If
End Sub

'------------------------------------------------
' B_6_Click
'------------------------------------------------
Private Sub B_6_Click()
    メッセージ = ""
    
    If B_6.Caption = "ログ一覧表" Then
        'Call ログ一覧表
    ElseIf B_6.Caption = "完全削除" Then
        Call 完全削除
    End If
End Sub

'------------------------------------------------
' フラグ解除
'------------------------------------------------
Private Sub フラグ解除()
'
    Dim wi01 As Integer
    Dim wsDB名 As String
    Dim ws01 As String, ws02 As String
'
    On Error GoTo フラグ解除_ERR
'
    '----------< DELETE_NEWREC >----------------------------------------------------
    If wNew企業名Key <> "" Then
        Call DELETE_NEWREC
    End If
'
    '----------< FLG 処理 >---------------------------------------------------------
    GRet = FLG_CHECK_NYURYOKUOFF(企業名Key)
'
    ws01 = "他端末で実行している" & GProduct & "を終了してください"
    ws02 = "フラグを解除します。よろしいですか？"
    GRet = MsgBox(ws01 + vbCrLf + vbCrLf + ws02, vbExclamation + vbOKCancel, "フラグ解除")
    If GRet = vbCancel Then
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        
        Exit Sub
    End If
'
    '----------< BUTTON_ENABLE_SET >------------------------------------------------
    Call BUTTON_ENABLE_SET(False)
'
    '----------< DAAA070_企業名マスタ >---------------------------------------------
    '----------< CHECK 入力中端末名 >-----------------------------------------------
    wstr = "Update"
    wstr = wstr + " DAAA070_企業名マスタ"
    wstr = wstr + " Set"
    wstr = wstr + " 入力中端末名 = ''"
    GDb2.Execute wstr
'
    ws01 = "": ws02 = ""
    wi01 = 0
    '----------< CHECK 処理中端末名 >-----------------------------------------------
    '----------< DAAA070_企業名マスタ >---------------------------------------------
    wstr = ""
    wstr = "Select * From DAAA070_企業名マスタ"
    wstr = wstr + " Where 処理中端末名 <> ''"
    Call AdoRecordsetOpen(GDb2, wRs2, wstr)
    wi01 = wRs2.RecordCount
    Do Until wRs2.eof
        '----------< AdoDbOpen_Check >----------------------------------------------
        wsDB名 = P8.FCStr(wRs2("DB名"))
        GRet = ADODBOPEN_CHECK("Jet", GDb, GSerDir + "\" + wsDB名, "", , GPwd, "排他")
        If GRet = True Then
            wRs2("処理中端末名") = ""
            wRs2.Update
        End If
        wRs2.MoveNext
    Loop
    wRs2.Close
    Set wRs2 = Nothing
'
    '----------< CHECK DAAA020_稼動中 >---------------------------------------------
    '----------< DAAA020_稼動中 >---------------------------------------------------
    If wi01 = 0 Then
        wstr = "Update"
        wstr = wstr + " DAAA020_稼動中"
        wstr = wstr + " Set"
        wstr = wstr + " 稼動中フラグ = 0"
        GDb2.Execute wstr
    Else
        wstr = "SELECT"
        wstr = wstr + " 稼動中.端末コンピュータ名,"
        wstr = wstr + " 稼動中.稼動中フラグ,"
        wstr = wstr + " 企業名マスタ.処理中端末名"
        wstr = wstr + " FROM DAAA020_稼動中 AS 稼動中"
        wstr = wstr + " LEFT JOIN DAAA070_企業名マスタ AS 企業名マスタ"
        wstr = wstr + " ON 稼動中.端末コンピュータ名 = 企業名マスタ.処理中端末名"
        wstr = wstr + " GROUP BY 稼動中.端末コンピュータ名,稼動中.稼動中フラグ,企業名マスタ.処理中端末名"
        Call AdoRecordsetOpen(GDb2, wRs2, wstr)
        Do Until wRs2.eof
            ws01 = P8.FCStr(wRs2("処理中端末名"))
            ws02 = P8.FCStr(wRs2("端末コンピュータ名"))
            
            wstr = "Update"
            wstr = wstr + " DAAA020_稼動中"
            wstr = wstr + " Set"
            
            If ws01 <> "" Then
                wstr = wstr + " 稼動中フラグ = 1"
            Else
                wstr = wstr + " 稼動中フラグ = 0"
            End If
            
            wstr = wstr + " Where 端末コンピュータ名 = '" & ws02 & "'"
            GDb2.Execute wstr
            
            wRs2.MoveNext
        Loop
        wRs2.Close
        Set wRs2 = Nothing
    End If
'
    '----------< LOG WRITE >--------------------------------------------------------
    GStr = "9," & "6," & P8.FCChr(GMyComputerName, 16) & "," & P8.FCChr("", 30) & "," & P8.FCChr("", 30)
    GRet = PUT_LOG_FILE(GStr)
'
    ' =========================================
    '                   後処理
    ' =========================================
    Call ADODC_RESET
'    Call BUTTON_ENABLE_SET(True)
    MsgBox "フラグを解除しました", vbInformation
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
フラグ解除_ERR:
    pERR_MES = pPROGRAM_ID + "/ フラグ解除() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' 復元処理
'------------------------------------------------
Private Sub 復元処理()
'
    Dim wDb As New ADODB.Connection
    Dim wRs3 As ADODB.Recordset
    
    Dim wretfn As String, wSDate As String
    Dim ws企業名Key As String, ws企業名 As String, wsDB名 As String
    Dim ws01 As String, ws02 As String
    Dim wi決算月 As Integer, wi決算締日 As Integer, wi回収有無 As Integer, wi支払有無 As Integer
    Dim wd実績年月 As Date
'
    On Error GoTo 復元処理_ERR
'
    '----------< DELETE_NEWREC >----------------------------------------------------
    If wNew企業名Key <> "" Then
        Call DELETE_NEWREC
    End If
'
    '----------< FLG 処理 >---------------------------------------------------------
    GRet = FLG_CHECK_NYURYOKUOFF(企業名Key)
'
    '----------< Check mdb >--------------------------------------------------------
    '----------< COMDLG >-----------------------------------------------------------
    wretfn = COMDLG("バックアップDBの復元", wBK, "AccessMdbファイル(*.mdb)|*.mdb", "")
    If wretfn = "" Then
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        
        Exit Sub
    ElseIf LCase(Dir(wretfn)) = LCase(GMain) Or LCase(Dir(wretfn)) = LCase(GTemp) Then
        MsgBox "システム用mdbと同名のmdbは復元対象に指定できません", vbCritical
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        
        Exit Sub
    End If
    
    GRet = CHECK_BKDIR(wretfn)
    If GRet <> True Then
        GRet = MsgBox("backup用フォルダを選択してください", vbCritical)
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        
        Exit Sub
    End If
'
    '----------< AdoDbOpen_Check >--------------------------------------------------
    GRet = ADODBOPEN_CHECK("Jet", wDb, wretfn, "", , GPwd, "排他")
    If GRet <> True Then
        GRet = MsgBox("他端末で使用中です" + vbCrLf + "使用可能になった時点で再度実行してください", vbExclamation + vbOKOnly)
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        '
        Exit Sub
    End If
'
    '----------< XXX.mdb Open >-----------------------------------------------------
    Call AdoDbOpen("Jet", wDb, wretfn, "", , GPwd)
    Set wRs3 = New ADODB.Recordset
    
    wstr = ""
    wstr = wstr + "Select "
    wstr = wstr + " count(Name) As Cnt "
    wstr = wstr + "From "
    wstr = wstr + "  MSysObjects "
    wstr = wstr + "Where "
    wstr = wstr + "      name = 'DAAA070_企業名マスタ'"
    wstr = wstr + "  and type = 1"
    On Error Resume Next
        wRs3.CursorType = adOpenKeyset
        wRs3.LockType = adLockOptimistic
        wRs3.Open wstr, wDb, , , adCmdText
    
        If Err.Number = "-2147217911" Then
            MsgBox "このMDBは" & GProduct & "用のMDBではありません", vbCritical
            wRs3.Close
            Set wRs3 = Nothing
            '----------< XXX.mdb Close >--------------------------------------------
            wDb.Close
            Set wDb = Nothing
            Err.Clear
            
            '----------< RESET Adodc >----------------------------------------------
            Call ADODC_RESET
            
            Exit Sub
        End If
    
    If wRs3("Cnt") = 0 Then
        MsgBox GProduct & "用MDBではありません", vbCritical
        wRs3.Close
        Set wRs3 = Nothing
        '----------< XXX.mdb Close >------------------------------------------------
        wDb.Close
        Set wDb = Nothing
        
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        
        Exit Sub
    End If
    wRs3.Close
    
    '----------< KXXX.mdb >---------------------------------------------------------
    '----------< Check Version >----------------------------------------------------
    '----------< RecordsetOpen DAAA000_バージョン >---------------------------------
    wstr = ""
    wstr = wstr + "Select * From DAAA000_バージョン"
    wstr = wstr + " Where System = 'System'"
    Call AdoRecordsetOpen(wDb, wRs3, wstr)
        If GVerNo <> wRs3("Version") Then
            MsgBox "バージョンの違うMDBを参照しています"
            wRs3.Close
            Set wRs3 = Nothing
            '----------< XXX.mdb Close >--------------------------------------------
            wDb.Close
            Set wDb = Nothing
            
            '----------< RESET Adodc >----------------------------------------------
            Call ADODC_RESET
            
            Exit Sub
        End If
    wRs3.Close
    Set wRs3 = Nothing
    
    '----------< DAAA070_企業名マスタ >---------------------------------------------
    ws企業名Key = ""
    wstr = ""
    wstr = wstr + "Select * "
    wstr = wstr + "From DAAA070_企業名マスタ "
    Call AdoRecordsetOpen(wDb, wRs3, wstr)
    If Not wRs3.eof Then
        ws企業名Key = P8.FCStr(wRs3("企業名key"))
    End If
    wRs3.Close
    Set wRs3 = Nothing
    '----------< XXX.mdb Close >----------------------------------------------------
    wDb.Close
    Set wDb = Nothing
'
    wSDate = Format(Now, "yyyy/mm/dd hh:nn:ss")
    '----------< FLG_KIGYOSHORI >---------------------------------------------------
    GRet = FLG_KIGYOSHORI(ws企業名Key, wSDate)
    If GRet <> True Then
        Exit Sub
    End If
    
    On Error GoTo 0
'

'
On Error GoTo 復元処理_ERR1
'
復元処理_ERR_RETRY1:
    '----------< List.mdb >--------------------------------------------------------
    '----------< DAAA070_企業名マスタ >---------------------------------------------
    ws01 = "": ws02 = "": wsDB名 = ""
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAA070_企業名マスタ"
    wstr = wstr + " Where"
    wstr = wstr + " 企業名Key = '" + ws企業名Key + "'"
    Call AdoRecordsetOpen(GDb2, wRs2, wstr)
    If Not wRs2.eof Then
        wsDB名 = P8.FCStr(wRs2("DB名"))
    End If
    wRs2.Close
    Set wRs2 = Nothing

    If wsDB名 <> "" Then
        '----------< AdoDbOpen_Check >----------------------------------------------
        GRet = ADODBOPEN_CHECK("Jet", GDb, GSerDir + "\" + wsDB名, "", , GPwd, "排他")
        If GRet <> True Then
            GRet = MsgBox("他端末で使用中です" + vbCrLf + "使用可能になった時点で再度実行してください", vbExclamation + vbOKOnly)
            ' =========================================
            '                   後処理
            ' =========================================
            Call ADODC_RESET
            MsgBox "復元作業ができませんでした", vbExclamation
        
            Exit Sub
        End If
        '
        GRet = MsgBox("同一企業名のデータが存在しています。データを上書きしてよろしいですか？" _
                , vbOKCancel + vbExclamation)
        If GRet = vbOK Then
            '----------< DeleteFile KXXX.mdb >------------------------------------------
            GRet = DeleteFile(GSerDir & "\" & wsDB名)
            If GRet = 0 Then
                GRet = MsgBox("他端末で使用中です" + vbCrLf + "使用可能になった時点で再度実行してください", vbExclamation + vbOKOnly)
                ' =========================================
                '                   後処理
                ' =========================================
                Call ADODC_RESET
                MsgBox "復元作業ができませんでした", vbExclamation
                
                Exit Sub
            End If
        Else
            ' =========================================
            '                   後処理
            ' =========================================
            Call ADODC_RESET
            MsgBox "復元作業ができませんでした", vbExclamation
            
            Exit Sub
        End If
    End If
    
    DoEvents
    
On Error GoTo 0
'

'
    On Error GoTo 復元処理_ERR
'
    '----------< BUTTON_ENABLE_SET >------------------------------------------------
    Call BUTTON_ENABLE_SET(False)
'
    '----------< Copy >-------------------------------------------------------------
    wsDB名 = ws企業名Key & ".mdb"
    GRet = FileMaker(wretfn, GSerDir & "\" & wsDB名)
    If GRet <> True Then
        GRet = MsgBox("他端末で使用中です" + vbCrLf + "使用可能になった時点で再度実行してください", vbExclamation + vbOKOnly)
        ' =========================================
        '                   後処理
        ' =========================================
        Call ADODC_RESET
'        Call BUTTON_ENABLE_SET(True)
        
        MsgBox "復元作業ができませんでした", vbExclamation
        
        Exit Sub
    End If
'
    '----------< AdoDbOpen_Check >--------------------------------------------------
    GRet = ADODBOPEN_CHECK("Jet", wDb, GSerDir + "\" + wsDB名, "", , GPwd, "排他")
    If GRet <> True Then
        GRet = MsgBox("他端末で使用中です" + vbCrLf + "使用可能になった時点で再度実行してください", vbExclamation + vbOKOnly)
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        '----------< BUTTON_ENABLE_SET >--------------------------------------------
'        Call BUTTON_ENABLE_SET(True)
        MsgBox "復元作業ができませんでした", vbExclamation
        '
        Exit Sub
    End If
    
    On Error GoTo 0
'

'
On Error GoTo 復元処理_ERR2
'
復元処理_ERR_RETRY2:
    '----------< KXXX.mdb Open >----------------------------------------------------
    Call AdoDbOpen("Jet", wDb, GSerDir + "\" + wsDB名, "", , GPwd)
    
    '----------< KXXX.mdb >---------------------------------------------------------
    '----------< DAAA010_基本情報 >-----------------------------------------
    wi決算月 = 3
    wi決算締日 = 31
    wi回収有無 = 0
    wi支払有無 = 0
    
    wstr = ""
    wstr = wstr + "Select 決算月,決算締日,回収有無,支払有無"
    wstr = wstr + " From DAAA010_基本情報"
    wstr = wstr + " Where System = 'System'"
    Call AdoRecordsetOpen(wDb, wRs3, wstr)
    If Not wRs3.eof Then
        wi決算月 = wRs3("決算月")
        wi決算締日 = wRs3("決算締日")
        wi回収有無 = wRs3("回収有無")
        wi支払有無 = wRs3("支払有無")
    End If
    
    wRs3.Close
    Set wRs3 = Nothing
    
    '----------< DAAA020_コントロール >-----------------------------------------
    wd実績年月 = CDate("2001/01/01")
    
    wstr = ""
    wstr = wstr + "Select 最終実績年月"
    wstr = wstr + " From DAAA020_コントロール"
    wstr = wstr + " Where System = 'System'"
    Call AdoRecordsetOpen(wDb, wRs3, wstr)
    If Not wRs3.eof Then
        wd実績年月 = Format(wRs3("最終実績年月"), "yyyy/mm/dd")
    End If
    
    wRs3.Close
    Set wRs3 = Nothing
        
    '----------< DAAA070_企業名マスタ >---------------------------------------------
    ws企業名 = ""
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAA070_企業名マスタ"
    wstr = wstr + " Where"
    wstr = wstr + " 企業名Key = '" + ws企業名Key + "'"
    Call AdoRecordsetOpen(wDb, wRs3, wstr)
    If Not wRs3.eof Then
        ws企業名 = P8.FCStr(wRs3("企業名"))
        wRs3("DB名") = wsDB名
        wRs3("復元日") = P8.FCDate(wSDate)
        
        wRs3.Update
        
        '----------< List.mdb >-----------------------------------------------------
        '----------< DAAA070_企業名マスタ >-----------------------------------------
        wstr = ""
        wstr = wstr + "Select * "
        wstr = wstr + " From DAAA070_企業名マスタ"
        wstr = wstr + " Where 企業名Key = '" + ws企業名Key + "'"
        Call AdoRecordsetOpen(GDb2, wRs2, wstr)
        If wRs2.eof Then
           wRs2.AddNew

           wRs2("企業名Key") = P8.FCStr(wRs3("企業名Key"))
        End If
            wRs2("企業名") = ws企業名
            wRs2("備考") = P8.FCStr(wRs3("備考"))
            wRs2("最新処理日") = P8.FCDate(wRs3("最新処理日"))
            wRs2("作成日") = P8.FCDate(wRs3("作成日"))
            wRs2("保存日") = P8.FCDate(wRs3("保存日"))
            wRs2("復元日") = P8.FCDate(wSDate)
            wRs2("削除日") = P8.FCDate(wRs3("削除日"))
            wRs2("DB名") = wsDB名
            wRs2("支店コード") = ""
            wRs2("支店名") = "単独企業"
            wRs2("親会社名") = ""
            wRs2("企業区分") = "単独企業"
            
            wRs2("更新回数") = P8.FCDbl(wRs3("更新回数"))
            wRs2("処理開始日付") = P8.FCDate(wRs3("処理開始日付"))
            wRs2("処理終了日付") = P8.FCDate(wRs3("処理終了日付"))
            wRs2("端末コンピュータ名") = P8.FCStr(wRs3("端末コンピュータ名"))
            wRs2("実績データ共有") = P8.FCDbl(wRs3("実績データ共有"))
            
            'V170 追加
            wRs2("決算月") = wi決算月
            wRs2("決算締日") = wi決算締日
            wRs2("回収有無") = wi回収有無
            wRs2("支払有無") = wi支払有無
            wRs2("最終実績年月") = Format(wd実績年月, "yyyy/mm/dd")
            
            wRs2.Update
        wRs2.Close
        Set wRs2 = Nothing
    
    End If
    
    wRs3.Close
    Set wRs3 = Nothing
    
    '----------< DAAA010_基本情報 >---------------------------------------------
    wstr = ""
    wstr = "UPDATE DAAA010_基本情報"
    wstr = wstr + " Set "
    wstr = wstr + " 支店コード = '',"
    wstr = wstr + " 支店名 = '単独企業',"
    wstr = wstr + " 企業区分 = '単独企業',"
    wstr = wstr + " 資金調達区分 = '2',"
    wstr = wstr + " 有担銀行 = 'ZZ',"
    wstr = wstr + " 無担銀行 = 'ZZ'"
    wstr = wstr & " WHERE System='System'"
    wDb.Execute wstr
    
    '----------< KXXX.mdb Close >---------------------------------------------------
    wDb.Close
    Set wDb = Nothing
    '
On Error GoTo 0
'

'
    On Error GoTo 復元処理_ERR
'
    ' =========================================
    '           KIGYOSHORI_END
    ' =========================================
    Call KIGYOSHORI_END(ws企業名Key, wSDate)
    
    '----------< KXXX.mdb Open >----------------------------------------------------
    Call AdoDbOpen("Jet", wDb, GSerDir + "\" + wsDB名, "", , GPwd)

    '----------< KXXX.mdb >---------------------------------------------------------
    '----------< DAAA070_企業名マスタ >---------------------------------------------
    wstr = "Update"
    wstr = wstr + " DAAA070_企業名マスタ"
    wstr = wstr + " Set"
    wstr = wstr + " 処理終了日付 = #" + wSDate + "#,"
    wstr = wstr + " 保存日 = #" + wSDate + "#,"
    wstr = wstr + " 入力中端末名 = '',"
    wstr = wstr + " 処理中端末名 = '',"
    wstr = wstr + " 端末コンピュータ名 = '" + GMyComputerName + "'"
    wstr = wstr + " Where 企業名Key = '" & ws企業名Key & "'"
    wDb.Execute wstr
    
    '----------< KXXX.mdb Close >---------------------------------------------------
    wDb.Close
    Set wDb = Nothing
'
    '----------< LOG WRITE >--------------------------------------------------------
    GStr = "9," & "4," & P8.FCChr(GMyComputerName, 16) & "," & P8.FCChr(ws企業名Key, 30) & "," & P8.FCChr(ws企業名, 30)
    GRet = PUT_LOG_FILE(GStr)
'
    ' =========================================
    '                   後処理
    ' =========================================
    Call ADODC_RESET
'    Call BUTTON_ENABLE_SET(True)
    MsgBox "企業DBの復元作業が完了しました", vbInformation
'
    Exit Sub
'
'----------< ERROR ROUTINE >--------------------------------------------------------
復元処理_ERR1:
    If Err.Number = -214727887 Or Err.Number = -2147217887 Then
        Sleep (1000)
        Resume 復元処理_ERR_RETRY1
    End If
    '
    pERR_MES = pPROGRAM_ID + "/ 復元処理() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                "他端末で同一企業(" & 企業名Key.Text & ")使用中です"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)
    '
    '----------< ADODC_RESET >------------------------------------------------------
    Call ADODC_RESET
    '
    Resume 復元処理_ERR_END
'
復元処理_ERR2:
    If Err.Number = -214727887 Or Err.Number = -2147217887 Then
        Sleep (1000)
        Resume 復元処理_ERR_RETRY2
    End If
    '
    pERR_MES = pPROGRAM_ID + "/ 復元処理() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                "他端末で同一企業(" & 企業名Key.Text & ")使用中です"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)
    '
    Sleep (1000)
    
    '----------< DeleteFile KXXX.mdb >----------------------------------------------
    GRet = DeleteFile(GSerDir & "\" & wsDB名)
    DoEvents

    Sleep (1000)
    
    '----------< Delete List data >-------------------------------------------------
    wstr = ""
    wstr = wstr + "Delete"
    wstr = wstr + " From DAAA070_企業名マスタ"
    wstr = wstr + " Where DB名 = '" + wsDB名 + "'"
    GDb2.Execute wstr
    DoEvents

    '----------< ADODC_RESET >------------------------------------------------------
    Call ADODC_RESET
'    Call BUTTON_ENABLE_SET(True)
    '
    Resume 復元処理_ERR_END
'
復元処理_ERR:
    pERR_MES = pPROGRAM_ID + "/ 復元処理() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
復元処理_ERR_END:
    Exit Sub
End Sub

'------------------------------------------------
' LISTDB処理
'------------------------------------------------
Private Sub LISTDB処理()
'
    Dim wDb As New ADODB.Connection
    Dim wRs3 As ADODB.Recordset
    
    Dim wSDate As String, wsDBKey As String, wsDB名 As String
    Dim ws01 As String
'
    '----------< DELETE_NEWREC >----------------------------------------------------
    If wNew企業名Key <> "" Then
        Call DELETE_NEWREC
    End If
'
    '----------< FLG 処理 >---------------------------------------------------------
    GRet = FLG_CHECK_NYURYOKUOFF(企業名Key)
'
    ws01 = "本支店管理を開始します"
    GRet = MsgBox(ws01 + vbCrLf + vbCrLf + "他端末で実行している" & GProduct & "を終了してください" + vbCrLf + "よろしいですか？", _
                         vbExclamation + vbOKCancel, "本支店管理")
    If GRet = vbCancel Then
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        
        Exit Sub
    End If
'
    '----------< BUTTON_ENABLE_SET >------------------------------------------------
    Call BUTTON_ENABLE_SET(False)
'
    wSDate = Format(Now, "yyyy/mm/dd hh:mm:ss")

    メッセージ = "しばらくお待ちください"
    メッセージ.Refresh
'
    '----------< FLG_KIGYOSHORI >---------------------------------------------------
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAA070_企業名マスタ "
    wstr = wstr + " Where 削除日 is null"
    Call AdoRecordsetOpen(GDb2, wRs3, wstr)
    If wRs3.eof Then
        MsgBox ("対象がありません")
        wRs3.Close
        Set wRs3 = Nothing
        
        ' =========================================
        '                   後処理
        ' =========================================
        Call ADODC_RESET
'        Call BUTTON_ENABLE_SET(True)

        メッセージ = ""
        
        Exit Sub
    Else
        Do While wRs3.eof = False
            wsDBKey = P8.FCStr(wRs3("企業名Key"))
            wsDB名 = P8.FCStr(wRs3("DB名"))
            
            If Dir(GSerDir + "\" + wsDB名) <> "" _
                And (P8.FCStr(wRs3("入力中端末名")) <> "" Or P8.FCStr(wRs3("処理中端末名")) <> "") Then
                    wRs3.Close
                    Set wRs3 = Nothing
                    '
                    GoTo LISTDB処理_ERR_CHECK_EXIT
            ElseIf Dir(GSerDir + "\" + wsDB名) <> "" _
                And (P8.FCStr(wRs3("入力中端末名")) = "" And P8.FCStr(wRs3("処理中端末名")) = "") Then
                
                '----------< FLG 処理 >---------------------------------------------
                GRet = FLG_KIGYOSHORI(wsDBKey, wSDate)
                If GRet <> True Then
                    メッセージ = ""
                    
                    Exit Sub
                End If
                
                '----------< AdoDbOpen_Check >--------------------------------------
                GRet = ADODBOPEN_CHECK("Jet", GDb, GSerDir + "\" + wsDB名, "", , GPwd, "排他")
                If GRet <> True Then
                    wRs3.Close
                    Set wRs3 = Nothing
                    '
                    GoTo LISTDB処理_ERR_CHECK_EXIT
                End If
            End If
            wRs3.MoveNext
        Loop
        wRs3.Close
        Set wRs3 = Nothing
    End If
'
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAA070_企業名マスタ "
    wstr = wstr + " Where 削除日 is null"
    Call AdoRecordsetOpen(GDb2, wRs3, wstr)
    Do Until wRs3.eof
        wsDBKey = P8.FCStr(wRs3("企業名Key"))
        wsDB名 = P8.FCStr(wRs3("DB名"))
        
        If Dir(GSerDir + "\" + wsDB名) = "" Then
            MsgBox "対象MDBが見つかりません。エクスプローラー側から削除、または移動された可能性があります｡ ", vbCritical
            
            GRet = MsgBox("対象をリストから抹消します", vbOKOnly + vbInformation)
            If GRet = vbOK Then
                '----------< Delete List data >-----------------------------------------
                wstr = ""
                wstr = wstr + "Delete"
                wstr = wstr + " From DAAA070_企業名マスタ"
                wstr = wstr + " Where DB名 = '" + wsDB名 + "'"
                GDb2.Execute (wstr)
            End If
        Else
            '----------< KXXX.mdb Open >----------------------------------------------------
            Call AdoDbOpen("Jet", wDb, GSerDir + "\" + wsDB名, "", , GPwd)
                wstr = "Update "
                wstr = wstr + "DAAA070_企業名マスタ"
                wstr = wstr + " Set"
                wstr = wstr + " 最新処理日 = #" + wSDate + "#,"
                wstr = wstr + " 処理開始日付 = #" + wSDate + "#,"
                wstr = wstr + " 処理終了日付 = #" + wSDate + "#,"
                wstr = wstr + " 端末コンピュータ名 = '" + GMyComputerName + "'"
                wstr = wstr + " Where 企業名Key = '" & wsDBKey & "'"
                wDb.Execute wstr
            '----------< KXXX.mdb Close >---------------------------------------------------
            wDb.Close
            Set wDb = Nothing
        End If
        '----------< LOG WRITE >--------------------------------------------------------
        GStr = "2," & "9," & P8.FCChr(GMyComputerName, 16) & "," & P8.FCChr(wsDBKey, 30) & "," & P8.FCChr(wsDBKey, 30)
        GRet = PUT_LOG_FILE(GStr)
        
        wRs3.MoveNext
    Loop
    wRs3.Close
    Set wRs3 = Nothing
'
    'FHA010_本支店管理メニュー.Show vbModal
'
    メッセージ = "しばらくお待ちください"
    メッセージ.Refresh
    
    '----------< LISTDB_KIGYOSHORI_END >--------------------------------------------
    Call LISTDB_KIGYOSHORI_END
    
    '----------< RESET Adodc >------------------------------------------------------
    Call ADODC_RESET
'    Call BUTTON_ENABLE_SET(True)
'
    メッセージ = ""
'
Exit Sub
'----------< ERROR ROUTINE >--------------------------------------------------------
LISTDB処理_ERR_CHECK_EXIT:
'
    GRet = MsgBox("他端末で使用中です" + vbCrLf + "使用可能になった時点で再度実行してください", vbExclamation + vbOKOnly)
'
    ' =========================================
    '           後処理
    ' =========================================
    Call ADODC_RESET    'Form_Loadでフラグ解除
'    Call BUTTON_ENABLE_SET(True)

    メッセージ = ""
'
End Sub

'------------------------------------------------
' LISTDB_KIGYOSHORI_END
'------------------------------------------------
Private Sub LISTDB_KIGYOSHORI_END()
'
    Dim wDb As New ADODB.Connection
    Dim wRs3 As ADODB.Recordset
    
    Dim wSDate As String, wsDBKey As String, wsDB名 As String
'
    On Error GoTo LISTDB_KIGYOSHORI_END_ERR
'
    wSDate = Format(Now, "yyyy/mm/dd hh:nn:ss")
'
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAA070_企業名マスタ "
    wstr = wstr + " Where 削除日 is null"
    Call AdoRecordsetOpen(GDb2, wRs3, wstr)
    Do Until wRs3.eof
        wsDBKey = P8.FCStr(wRs3("企業名Key"))
        wsDB名 = P8.FCStr(wRs3("DB名"))
        
        ' =========================================
        '           KIGYOSHORI_END
        ' =========================================
        Call KIGYOSHORI_END(wsDBKey, wSDate)
        
        '----------< KXXX.mdb Open >------------------------------------------------
        Call AdoDbOpen("Jet", wDb, GSerDir + "\" + wsDB名, "", , GPwd)
    
        '----------< KXXX.mdb >-----------------------------------------------------
        '----------< DAAA070_企業名マスタ >-----------------------------------------
        wstr = "Update"
        wstr = wstr + " DAAA070_企業名マスタ"
        wstr = wstr + " Set"
        wstr = wstr + " 処理終了日付 = #" + wSDate + "#,"
        wstr = wstr + " 保存日 = #" + wSDate + "#,"
        wstr = wstr + " 入力中端末名 = '',"
        wstr = wstr + " 処理中端末名 = ''"
        wstr = wstr + " Where 企業名Key = '" & wsDBKey & "'"
        wDb.Execute wstr
    
        '----------< KXXX.mdb Close >-----------------------------------------------
        wDb.Close
        Set wDb = Nothing
        '
        '----------< LOG WRITE >----------------------------------------------------
        GStr = "3," & "9," & P8.FCChr(GMyComputerName, 16) & "," & P8.FCChr(wsDBKey, 30) & "," & P8.FCChr(wsDBKey, 30)
        GRet = PUT_LOG_FILE(GStr)
'
        wRs3.MoveNext
    Loop
    
    wRs3.Close
    Set wRs3 = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
LISTDB_KIGYOSHORI_END_ERR:
    pERR_MES = pPROGRAM_ID + "/ LISTDB_KIGYOSHORI_END() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' 最適化
'------------------------------------------------
Private Sub 最適化()
'
    Dim wDb As New ADODB.Connection
    
    Dim wretfn As String, wSDate As String
    Dim ws企業名Key As String
    Dim wsCDB名 As String, wsDB名 As String
    Dim ws01 As String
'
    On Error GoTo 最適化_ERR
'
    '----------< DELETE_NEWREC >----------------------------------------------------
    If wNew企業名Key <> "" Then
        Call DELETE_NEWREC
    End If
'
    If 新規変更.Caption <> "選択／内容変更" Then
        MsgBox "対象が選択されていません"
        '----------< FLG 処理 >-----------------------------------------------------
        GRet = FLG_CHECK_NYURYOKUOFF(企業名Key)
'
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        
        Exit Sub
    End If
'
    wsCDB名 = wDB名
    If Dir(GSerDir + "\" + wsCDB名) = "" Then
        MsgBox "対象MDBが見つかりません。エクスプローラー側から削除、または移動された可能性があります｡ ", vbCritical
        
        GRet = MsgBox("対象をリストから抹消しますか？", vbYesNo + vbExclamation)
        If GRet = vbYes Then
            '----------< BUTTON_ENABLE_SET >----------------------------------------
            Call BUTTON_ENABLE_SET(False)
            
            '----------< Delete List data >-----------------------------------------
            wstr = ""
            wstr = wstr + "Delete"
            wstr = wstr + " From DAAA070_企業名マスタ"
            wstr = wstr + " Where DB名 = '" + wsCDB名 + "'"
            GDb2.Execute (wstr)
            
            ' =========================================
            '                   後処理
            ' =========================================
            Call ADODC_RESET
'            Call BUTTON_ENABLE_SET(True)
            
            Exit Sub
        Else
            '----------< FLG 処理 >-------------------------------------------------
            GRet = FLG_CHECK_NYURYOKUOFF(企業名Key)
            '----------< RESET Adodc >----------------------------------------------
            Call ADODC_RESET
            
            Exit Sub
        End If
    End If
'
    wSDate = Format(Now, "yyyy/mm/dd hh:nn:ss")
    
    '----------< FLG_KIGYOSHORI >---------------------------------------------------
    GRet = FLG_KIGYOSHORI(企業名Key, wSDate)
    If GRet <> True Then
        Exit Sub
    End If
'
    '----------< BUTTON_ENABLE_SET >------------------------------------------------
    Call BUTTON_ENABLE_SET(False)
'
    ' =========================================
    '           KIGYOSHORI_END
    ' =========================================
    Call KIGYOSHORI_END(ws企業名Key, wSDate)
'
    '----------< 最適化 >-----------------------------------------------------------
    GRet = MXA030_CompactDb(wsCDB名)
    If GRet <> True Then
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        
        '----------< BUTTON_ENABLE_SET >--------------------------------------------
        Call BUTTON_ENABLE_SET(True)
        MsgBox "企業DBの最適化ができませんでした", vbExclamation
                
        Exit Sub
    End If
'
    ' =========================================
    '                   後処理
    ' =========================================
    Call ADODC_RESET
'    Call BUTTON_ENABLE_SET(True)
    MsgBox "企業DBの最適化が完了しました", vbInformation
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
最適化_ERR:
    pERR_MES = pPROGRAM_ID + "/ 最適化() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' サーバー指定
'------------------------------------------------
Private Sub サーバー指定()
'
    Dim wDb As New ADODB.Connection
    Dim objDriveSystem As Object
    Dim objDrive As Object
    
    Dim wl01 As Long
    Dim j As Integer
    Dim FLG_NET As Boolean
    Dim strDrive As String, wsRet As String
    Dim ws01 As String, ws02 As String
'
    On Error GoTo サーバー指定_ERR
'
    FLG_NET = False
'
    '----------< DELETE_NEWREC >----------------------------------------------------
    If wNew企業名Key <> "" Then
        Call DELETE_NEWREC
    End If
'
    '----------< FLG 処理 >---------------------------------------------------------
    GRet = FLG_CHECK_NYURYOKUOFF(企業名Key)
'
    If GSys.Lan = False Then
        GSerDir = GCurDir
        GSerComputerName = GMyComputerName
        
        Exit Sub
    End If
'
    GRet = MsgBox("他端末をサーバーに指定します。" + vbCrLf + "よろしいですか？", vbExclamation + vbYesNo, "サーバー指定")
    If GRet = vbNo Then
'        If GCurDir <> GSerDir Then
            ws01 = "サーバーを自端末(" & GMyComputerName & ")にリセットしますか？"
            GRet = MsgBox(ws01, vbExclamation + vbYesNo, "サーバー指定")
            If GRet = vbYes Then
                GSerDir = GCurDir
                GSerComputerName = GMyComputerName
            
                '----------< BUTTON_ENABLE_SET >------------------------------------
                Call BUTTON_ENABLE_SET(False)
            
                '----------< K000.mdb Open >----------------------------------------
                Call AdoDbOpen("Jet", wDb, GCurDir + "\" + GTemp, "", , GPwd)
            
                wstr = "UPDATE DAAA020_コントロール"
                wstr = wstr + " Set サーバーフォルダ = ''"
                wstr = wstr + " Where System = 'System'"
                wDb.Execute wstr
            
                '----------< K000.mdb Close >---------------------------------------
                wDb.Close
                Set wDb = Nothing
                
                ' =========================================
                '                   後処理
                ' =========================================
                Call ADODC_RESET
'                Call BUTTON_ENABLE_SET(True)
                
                Exit Sub
            Else
                '----------< RESET Adodc >------------------------------------------
                Call ADODC_RESET

                Exit Sub
            End If
'        Else
'            ' =========================================
'            '                   後処理
'            ' =========================================
'            Call ADODC_RESET
'
'            Exit Sub
'        End If
    End If
'
    '----------< BrowseFolder >-----------------------------------------------------
    wsRet = BrowseFolder(GProduct & "フォルダを選択してください")
    If wsRet = "" Then
        
    '----------< InputBox >-----------------------------------------------------
        wsRet = InputBox("借換たろうフォルダをフルパスで入力してください。", "借換たろうフォルダ入力", "\借換たろう")
            If wsRet = "" Then
            '----------< RESET Adodc >--------------------------------------------------
            Call ADODC_RESET
            
            Exit Sub
        End If
    End If
'
    '----------< CHECK 金剛石フォルダ >---------------------------------------------
    GRet = CHECK_KDIR(wsRet)
    If GRet <> True Then
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        
        MsgBox "サーバー指定できませんでした", vbExclamation
        Exit Sub
    End If
'
    '----------< BUTTON_ENABLE_SET >------------------------------------------------
    Call BUTTON_ENABLE_SET(False)
'
    ws01 = "": ws02 = ""
    If Left$(wsRet, 2) Like "*:" Then
        strDrive = Left$(wsRet, 1)
        Set objDriveSystem = CreateObject("Scripting.FileSystemObject")
        Set objDrive = objDriveSystem.GetDrive(strDrive)
      
        If objDrive.DriveType = 3 Then    'ネットワークドライブ
            FLG_NET = True
            
            wl01 = Len(wsRet)
            If Mid$(wsRet, 3, 1) <> "\" Then
                ws01 = "\" & Mid$(wsRet, 3, wl01)
            Else
                ws01 = Mid$(wsRet, 3, wl01)
            End If
        
            wsRet = objDrive.ShareName & ws01
        End If
      
        Set objDriveSystem = Nothing
        Set objDrive = Nothing
        
    ElseIf Left$(wsRet, 2) = "\\" Then
        FLG_NET = True
    Else
        ' =========================================
        '                   後処理
        ' =========================================
        Call ADODC_RESET
'        Call BUTTON_ENABLE_SET(True)
        MsgBox "サーバー指定できませんでした", vbExclamation
        
        Exit Sub
    End If
    
    ws01 = "": ws02 = ""
    If FLG_NET = True Then
        wl01 = Len(wsRet)
        For j = 3 To wl01
            ws01 = Mid$(wsRet, j, 1)
            If ws01 <> "\" Then
                ws02 = ws02 & ws01
            Else
                Exit For
            End If
        Next j
        
        GSerDir = wsRet
        GSerComputerName = ws02
    Else
        GRet = MsgBox("他端末を選択してください", vbExclamation + vbOKOnly)
        ' =========================================
        '                   後処理
        ' =========================================
        Call ADODC_RESET
'        Call BUTTON_ENABLE_SET(True)
        MsgBox "サーバー指定できませんでした", vbExclamation
        
        Exit Sub
    End If
    
    If LCase(GSerComputerName) = LCase(GMyComputerName) Then
        GSerDir = GCurDir
        GSerComputerName = GMyComputerName
    
        GRet = MsgBox("他端末を選択してください", vbExclamation + vbOKOnly)
        ' =========================================
        '                   後処理
        ' =========================================
        Call ADODC_RESET
'        Call BUTTON_ENABLE_SET(True)
        MsgBox "サーバー指定できませんでした", vbExclamation
        
        Exit Sub
    End If
'
    メッセージ = "サーバーをセットしています。しばらくお待ちください。"
'
    '----------< K000.mdb Open >----------------------------------------------------
    Call AdoDbOpen("Jet", wDb, GCurDir + "\" + GTemp, "", , GPwd)

        wstr = "UPDATE DAAA020_コントロール"
        wstr = wstr + " Set サーバーフォルダ = '" & GSerDir & "'"
        wstr = wstr + " Where System = 'System'"
        wDb.Execute wstr

    '----------< K000.mdb Close >---------------------------------------------------
    wDb.Close
    Set wDb = Nothing
'
    '----------< List.mdb Close >---------------------------------------------------
    GDb2.Close
    Set GDb2 = Nothing
'
    ' =========================================
    '           CHECK SET SerVer
    ' =========================================
    GRet = MAA100_MDBVER("指定")
    If GRet <> True Then
        '----------< K000.mdb Open >------------------------------------------------
        Call AdoDbOpen("Jet", wDb, GCurDir + "\" + GTemp, "", , GPwd)

            wstr = ""
            wstr = "UPDATE DAAA020_コントロール"
            wstr = wstr + " SET サーバーフォルダ ='" & GCurDir & "'"
            wstr = wstr + " Where System = 'System'"
            wDb.Execute wstr
    
        '----------< K000.mdb Close >-----------------------------------------------
        wDb.Close
        Set wDb = Nothing
        
        GSerDir = GCurDir
        GSerComputerName = GMyComputerName
    
        ws01 = "サーバー指定できませんでした"
    Else
        ws01 = "サーバーをセットしました"
    End If
'
    DoEvents
    Set DataGrid1.DataSource = Nothing
    Adodc1.Recordset.Close
'
    FLG_DIR = False
    FLG_KIGYO = False
    DoEvents
'
    Call Form_Load
'
    DoEvents
'
    '----------< BUTTON_ENABLE_SET >------------------------------------------------
    Call BUTTON_ENABLE_SET(True)
        メッセージ = ws01
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
サーバー指定_ERR:
    pERR_MES = pPROGRAM_ID + "/ サーバー指定() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' CSV出力先指定
'------------------------------------------------
Private Sub CSV出力先指定()
'
    Dim objDriveSystem As Object
    Dim objDrive As Object
    
    Dim wDb As New ADODB.Connection
    
    Dim wl01 As Long
    Dim j As Integer
    Dim FLG_NET As Boolean
    Dim strDrive As String, wsRet As String
    Dim ws01 As String, ws02 As String
'
    On Error GoTo CSV出力先指定_ERR
'
    FLG_NET = False
'
    '----------< DELETE_NEWREC >----------------------------------------------------
    If wNew企業名Key <> "" Then
        Call DELETE_NEWREC
    End If
'
    '----------< FLG 処理 >---------------------------------------------------------
    GRet = FLG_CHECK_NYURYOKUOFF(企業名Key)
'
    If GSys.Lan = False Then
        GSerDir = GCurDir
        GSerComputerName = GMyComputerName
        
        Exit Sub
    End If
'
    '----------< BrowseFolder >-----------------------------------------------------
    If GCsvPath = "" Then
        wsRet = BrowseFolder("借換たろうCSV出力先指定フォルダを選択してください。")
    Else
        ws01 = "   (登録CSV出力先:" + GCsvPath + ")"
        wsRet = BrowseFolder("借換たろうCSV出力先指定フォルダを選択してください。" + vbCrLf + ws01)
    End If
    If wsRet = "" Then
        
    '----------< InputBox >-----------------------------------------------------
        wsRet = InputBox("CSV出力先指定フォルダをフルパスで入力してください。", "借換たろうCSV出力先指定フォルダ入力", "\借換たろうCSV\")
        If wsRet = "" Then
            '----------< RESET Adodc >--------------------------------------------------
            Call ADODC_RESET
            
            Exit Sub
        End If
    End If
'
    '----------< BUTTON_ENABLE_SET >------------------------------------------------
    Call BUTTON_ENABLE_SET(False)
'
    ws01 = "": ws02 = ""
    If Left$(wsRet, 2) Like "*:" Then
        strDrive = Left$(wsRet, 1)
        Set objDriveSystem = CreateObject("Scripting.FileSystemObject")
        Set objDrive = objDriveSystem.GetDrive(strDrive)
      
        If objDrive.DriveType = 3 Then    'ネットワークドライブ
            FLG_NET = True
            
            wl01 = Len(wsRet)
            If Mid$(wsRet, 3, 1) <> "\" Then
                ws01 = "\" & Mid$(wsRet, 3, wl01)
            Else
                ws01 = Mid$(wsRet, 3, wl01)
            End If
        
            wsRet = objDrive.ShareName & ws01
        End If
      
        Set objDriveSystem = Nothing
        Set objDrive = Nothing
        
    ElseIf Left$(wsRet, 2) = "\\" Then
        FLG_NET = True
    End If
    
    If wsRet <> "" Then
        If Right$(wsRet, 1) <> "\" Then
            wsRet = wsRet & "\"
        End If
        ' =========================================
        '                   UPDATE
        ' =========================================
        '----------< K000.mdb Open >----------------------------------------
        Call AdoDbOpen("Jet", wDb, GCurDir + "\" + GMain, "", , GPwd)
            wstr = ""
            wstr = "UPDATE LIST000_データ保存先マスタ"
            wstr = wstr + " SET CSVPATH ='" & wsRet & "'"
            wstr = wstr + " Where System = 'System'"
            wDb.Execute wstr

        '----------< K000.mdb Close >---------------------------------------
        wDb.Close
        Set wDb = Nothing
        
        GCsvPath = wsRet
        
        DoEvents
    
        ' =========================================
        '           Csv File Drive
        ' =========================================
        Call MX040_CsvPath
        
        MsgBox "CSV出力先指定を登録しました。", vbInformation
        
    Else
        
        ' =========================================
        '                   後処理
        ' =========================================
        Call ADODC_RESET
'        Call BUTTON_ENABLE_SET(True)
        MsgBox "CSV出力先指定できませんでした", vbExclamation
        
        Exit Sub
    End If
'
    '----------< List.mdb Close >---------------------------------------------------
    GDb2.Close
    Set GDb2 = Nothing
'
    DoEvents
    Set DataGrid1.DataSource = Nothing
    Adodc1.Recordset.Close
'
    FLG_DIR = False
    FLG_KIGYO = False
    DoEvents
'
    Call Form_Load
'
    DoEvents
'
    '----------< BUTTON_ENABLE_SET >------------------------------------------------
    Call BUTTON_ENABLE_SET(True)
        メッセージ = ws01
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
CSV出力先指定_ERR:
    pERR_MES = pPROGRAM_ID + "/ CSV出力先指定() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' CHECK_KDIR
'------------------------------------------------
Private Function CHECK_KDIR(pPath As String) As Boolean
'
    Dim wi01 As Integer
    Dim j As Integer
    Dim ws01 As String, ws02 As String
'
    CHECK_KDIR = False
'
    ws01 = "": ws02 = ""
    wi01 = Len(pPath)
    For j = wi01 To 1 Step -1
        ws01 = Mid$(pPath, j, 1)
        If ws01 <> "\" Then
            ws02 = ws01 & ws02
        Else
            Exit For
        End If
    Next j
            
    '金剛石 or 借換たろう！
    If GProduct <> "金剛石" Then
        If ws02 <> "借換たろう" Then
            GRet = MsgBox("借換たろうフォルダを選択して下さい", vbOKOnly + vbCritical)
            Exit Function
        End If
    Else
        If ws02 <> "金剛石" Then
            GRet = MsgBox("金剛石フォルダを選択して下さい", vbOKOnly + vbCritical)
            Exit Function
        End If
    End If
'
    CHECK_KDIR = True
'
End Function

'------------------------------------------------
' CHECK_BKDIR
'------------------------------------------------
Private Function CHECK_BKDIR(pPath As String) As Boolean
'
    Dim objDriveSystem As Object
    Dim objDrive As Object
    
    Dim wl01 As Long
    Dim ws01 As String, ws02 As String
'
    CHECK_BKDIR = False
'
    ws01 = "": ws02 = ""
    If Left$(pPath, 2) Like "*:" Then
        ws01 = Left$(pPath, 1)
        Set objDriveSystem = CreateObject("Scripting.FileSystemObject")
        Set objDrive = objDriveSystem.GetDrive(ws01)
      
        If objDrive.DriveType = 3 Then    'ネットワークドライブ
            Exit Function
            
            Set objDriveSystem = Nothing
            Set objDrive = Nothing
        End If
      
        Set objDriveSystem = Nothing
        Set objDrive = Nothing
        
    ElseIf Left$(pPath, 2) = "\\" Then
            Exit Function
    End If

    ws01 = GetFullPathTOpathOnly(pPath)
    ws02 = Dir(ws01 & GTemp)
    If ws02 <> "" Then
        Exit Function
    End If
'
    CHECK_BKDIR = True
'
End Function

'------------------------------------------------
' GetFullPathTOpathOnly
'------------------------------------------------
Private Function GetFullPathTOpathOnly(ByVal strFullPathFileName As String) As String
'
   Dim I As Integer 'ループカウンタ
   Dim FNSize As Long '(ファイルが指定された)フルパスの文字サイズ
   Dim strTMP As String '作業用
   Dim s As String 'フルパス
   'フルパスをｓへ格納
   s = strFullPathFileName
   '(ファイルが指定された)フルパスからサイズを取得
   FNSize = Len(s)
   '親か子フォルダかを判定
   If FNSize <= 3 Then
     '"c:" のみだった場合は、逆に"\" を付ける
     If Right(s, 1) <> "\" Then s = s & "\"
   Else
      '"\”を後ろから探す
     For I = 1 To FNSize
      strTMP = Mid(s, Len(s) - I, 1)
       '\を見つけた？
       If strTMP = "\" Then
          'ファイル名を除いたパスを返す
         GetFullPathTOpathOnly = StrConv(Left(s, Len(s) - I), vbLowerCase)
          Exit For
      End If
    Next I
  End If
End Function

'------------------------------------------------
' バックアップ
'------------------------------------------------
Private Sub バックアップ()
'
    Dim wDb As New ADODB.Connection
    
    Dim wretfn As String, wSDate As String
    Dim ws企業名Key As String
    Dim wsCDB名 As String, wsDB名 As String
    Dim ws01 As String
'
    On Error GoTo バックアップ_ERR
'
    '----------< DELETE_NEWREC >----------------------------------------------------
    If wNew企業名Key <> "" Then
        Call DELETE_NEWREC
    End If
'
    If 新規変更.Caption <> "選択／内容変更" Then
        MsgBox "対象が選択されていません"
        '----------< FLG 処理 >-----------------------------------------------------
        GRet = FLG_CHECK_NYURYOKUOFF(企業名Key)
'
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        
        Exit Sub
    End If
'
    wsCDB名 = wDB名
    If Dir(GSerDir + "\" + wsCDB名) = "" Then
        MsgBox "対象MDBが見つかりません。エクスプローラー側から削除、または移動された可能性があります｡ ", vbCritical
        
        GRet = MsgBox("対象をリストから抹消しますか？", vbYesNo + vbExclamation)
        If GRet = vbYes Then
            '----------< BUTTON_ENABLE_SET >----------------------------------------
            Call BUTTON_ENABLE_SET(False)
            
            '----------< Delete List data >-----------------------------------------
            wstr = ""
            wstr = wstr + "Delete"
            wstr = wstr + " From DAAA070_企業名マスタ"
            wstr = wstr + " Where DB名 = '" + wsCDB名 + "'"
            GDb2.Execute (wstr)
            
            ' =========================================
            '                   後処理
            ' =========================================
            Call ADODC_RESET
'            Call BUTTON_ENABLE_SET(True)
            
            Exit Sub
        Else
            '----------< FLG 処理 >-------------------------------------------------
            GRet = FLG_CHECK_NYURYOKUOFF(企業名Key)
            '----------< RESET Adodc >----------------------------------------------
            Call ADODC_RESET
            
            Exit Sub
        End If
    End If
'
    wSDate = Format(Now, "yyyy/mm/dd hh:nn:ss")
    '----------< FLG_KIGYOSHORI >---------------------------------------------------
    GRet = FLG_KIGYOSHORI(企業名Key, wSDate)
    If GRet <> True Then
        Exit Sub
    End If
'
    '----------< LOG WRITE >--------------------------------------------------------
    GStr = "9," & "5," & P8.FCChr(GMyComputerName, 16) & "," & P8.FCChr(企業名Key, 30) & "," & P8.FCChr(企業名, 30)
    GRet = PUT_LOG_FILE(GStr)
'
    DoEvents
'
    ws01 = ""
    ws企業名Key = ""
    ws企業名Key = 企業名Key.Text
    wsDB名 = "backup" & ws企業名Key & ".mdb"
    '----------< COMDLG >-----------------------------------------------------------
    wretfn = COMDLG("対象企業のバックアップ", wBK, "AccessMdbファイル(*.mdb)|*.mdb", wsDB名)
    If wretfn = "" Then
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        
        Exit Sub
    End If
    
    GRet = CHECK_BKDIR(wretfn)
    If GRet <> True Then
        GRet = MsgBox("backup用フォルダを選択してください", vbCritical)
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        
        Exit Sub
    End If
    
    If Dir(wretfn) <> "" Then
        If LCase(Dir(wretfn)) = "list.mdb" Or LCase(Dir(wretfn)) = "k000.mdb" Then
            GRet = MsgBox("システム用mdbと同名のmdbには上書き、または保存できません", vbCritical)
            '----------< RESET Adodc >----------------------------------------------
            Call ADODC_RESET
        
            Exit Sub
        End If
            
        GRet = MsgBox("ファイルが存在します。上書きしますか？", vbOKCancel + vbExclamation, "バックアップ")
        If GRet = vbOK Then
            '----------< DeleteFile >-----------------------------------------------
            GRet = DeleteFile(wretfn)
            If GRet = 0 Then
                GRet = MsgBox("他端末で使用中です" + vbCrLf + "使用可能になった時点で再度実行してください", vbExclamation + vbOKOnly)
                '----------< RESET Adodc >------------------------------------------
                Call ADODC_RESET
                MsgBox "企業DBの保存ができませんでした", vbExclamation
        
                Exit Sub
            End If
        Else
            '----------< RESET Adodc >----------------------------------------------
            Call ADODC_RESET
        
            Exit Sub
        End If
    End If
    DoEvents
'
    '----------< AdoDbOpen_Check >----------------------------------------------
    GRet = ADODBOPEN_CHECK("Jet", GDb, GSerDir + "\" + wsCDB名, "", , GPwd, "排他")
    If GRet <> True Then
        wRs2.Close
        Set wRs2 = Nothing
        
        GRet = MsgBox("他端末で使用中です" + vbCrLf + "使用可能になった時点で再度実行してください", vbExclamation + vbOKOnly)
        
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        MsgBox "企業DBの保存ができませんでした", vbExclamation
            
        Exit Sub
    End If
'
    '----------< List.mdb >---------------------------------------------------------
    '----------< DAAA070_企業名マスタ >---------------------------------------------
    wstr = ""
    wstr = wstr + "Update  DAAA070_企業名マスタ"
    wstr = wstr + " Set"
    wstr = wstr + " 保存日 = #" + wSDate + "#,"
    wstr = wstr + " 端末コンピュータ名 = '" + GMyComputerName + "'"
    wstr = wstr + " Where 企業名Key = '" + 企業名Key.Text + "'"
    wstr = wstr + " And DB名 = '" + wsCDB名 + "'"
    GDb2.Execute wstr
'
    '----------< BUTTON_ENABLE_SET >------------------------------------------------
    Call BUTTON_ENABLE_SET(False)
'
    ' =========================================
    '           KIGYOSHORI_END
    ' =========================================
    Call KIGYOSHORI_END(ws企業名Key, wSDate)
    
    GRet = ADODBOPEN_CHECK("Jet", wDb, GSerDir + "\" + wsCDB名, "", , GPwd, "排他")
    If GRet <> True Then
        GRet = MsgBox("他端末で使用中です" + vbCrLf + "使用可能になった時点で再度実行してください", vbExclamation + vbOKOnly)
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        '----------< BUTTON_ENABLE_SET >--------------------------------------------
'        Call BUTTON_ENABLE_SET(True)
        MsgBox "企業DBの保存ができませんでした", vbExclamation
        '
        Exit Sub
    End If
    
    '----------< KXXX.mdb Open >----------------------------------------------------
    Call AdoDbOpen("Jet", wDb, GSerDir + "\" + wsCDB名, "", , GPwd)
    
    '----------< KXXX.mdb >---------------------------------------------------------
    '----------< DAAA070_企業名マスタ >---------------------------------------------
    ws01 = ws企業名Key & ".mdb"
    
    wstr = "Update"
    wstr = wstr + " DAAA070_企業名マスタ"
    wstr = wstr + " Set"
    wstr = wstr + " 処理終了日付 = #" + wSDate + "#,"
    wstr = wstr + " 保存日 = #" + wSDate + "#,"
    wstr = wstr + " DB名 = '" & ws01 & "',"
    wstr = wstr + " 入力中端末名 = '',"
    wstr = wstr + " 処理中端末名 = '',"
    wstr = wstr + " 端末コンピュータ名 = '" + GMyComputerName + "'"
    wstr = wstr + " Where 企業名Key = '" & ws企業名Key & "'"
    wDb.Execute wstr
    
    '----------< KXXX.mdb Close >---------------------------------------------------
    wDb.Close
    Set wDb = Nothing
'
    '----------< Copy >-------------------------------------------------------------
    GRet = FileMaker(GSerDir & "\" & wsCDB名, wretfn)
    If GRet <> True Then
        '----------< BUTTON_ENABLE_SET >--------------------------------------------
        Call BUTTON_ENABLE_SET(True)
        MsgBox "企業DBの保存ができませんでした", vbExclamation
        
        Exit Sub
    End If
'
    ' =========================================
    '                   後処理
    ' =========================================
    Call ADODC_RESET
'    Call BUTTON_ENABLE_SET(True)
    MsgBox "企業DBの保存作業が完了しました", vbInformation
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
バックアップ_ERR:
    pERR_MES = pPROGRAM_ID + "/ バックアップ() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' システムバックアップ
'------------------------------------------------
Private Sub システムバックアップ()
'
    Dim j As Integer
    Dim wi01 As Integer
    Dim ws01 As String, ws02 As String
    Dim wsRet As String, wGDir As String
'
    On Error GoTo システムバックアップ_ERR
'
    '----------< DELETE_NEWREC >----------------------------------------------------
    If wNew企業名Key <> "" Then
        Call DELETE_NEWREC
    End If
'
    '----------< FLG 処理 >---------------------------------------------------------
    GRet = FLG_CHECK_NYURYOKUOFF(企業名Key)
'
    wGDir = ""
    If GSys.Lan = True Then
        If GSerDir <> CurDir Then
            GRet = MsgBox("他端末の" & GProduct & "フォルダ(" & GSerDir & ") " + vbCrLf + "をバックアップします。  よろしいですか？", _
                    vbYesNo + vbExclamation, GProduct & "フォルダの保存")
            If GRet = vbNo Then
                GRet = MsgBox("自端末の" & GProduct & "フォルダ(" & GCurDir & ") " + vbCrLf + "をバックアップします。  よろしいですか？", _
                        vbYesNo + vbExclamation, GProduct & "フォルダの保存")
                If GRet = vbNo Then
                    '----------< RESET Adodc >--------------------------------------
                    Call ADODC_RESET
                    
                    Exit Sub
                End If
                wGDir = GCurDir
            Else
                wGDir = GSerDir
            End If
        Else
            GRet = MsgBox("自端末の" & GProduct & "フォルダ(" & GCurDir & ") " + vbCrLf + "をバックアップします。  よろしいですか？", _
                    vbYesNo + vbExclamation, GProduct & "フォルダの保存")
            If GRet = vbNo Then
                '----------< RESET Adodc >------------------------------------------
                Call ADODC_RESET
                
                Exit Sub
            End If
                
            wGDir = GCurDir
        End If
    Else
        GRet = MsgBox(GProduct & "フォルダ(" & GCurDir & ") " + vbCrLf + "をバックアップします。  よろしいですか？", _
                vbYesNo + vbExclamation, GProduct & "フォルダの保存")
        If GRet = vbNo Then
            '----------< RESET Adodc >----------------------------------------------
            Call ADODC_RESET
            
            Exit Sub
        End If
        wGDir = GCurDir
    End If
    
    '----------< BrowseFolder >-----------------------------------------------------
    wsRet = BrowseFolder(GProduct & "フォルダのバックアップ先を選択してください" + vbCrLf + GProduct & "フォルダ以外を選択してください")
    If wsRet = "" Then
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        
        Exit Sub
    End If
'
    If GCurDir = wsRet Or GSerDir = wsRet Then
        GRet = MsgBox(GProduct & "フォルダ以外のフォルダを選択してください", vbOKOnly + vbExclamation, GProduct & "フォルダの保存")
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        
        Exit Sub
    End If
    
    If Len(wsRet) >= Len(GCurDir) Then
        wi01 = Len(GCurDir)
        If Left$(wsRet, wi01) = GCurDir Then
            GRet = MsgBox(GProduct & "フォルダ以外のフォルダを選択してください", vbOKOnly + vbExclamation, GProduct & "フォルダの保存")
            '----------< RESET Adodc >----------------------------------------------
            Call ADODC_RESET
            
            Exit Sub
        End If
    End If
    
    If Len(wsRet) >= Len(GSerDir) Then
        wi01 = Len(GSerDir)
        If Left$(wsRet, wi01) = GSerDir Then
            GRet = MsgBox(GProduct & "フォルダ以外のフォルダを選択してください", vbOKOnly + vbExclamation, GProduct & "フォルダの保存")
            '----------< RESET Adodc >----------------------------------------------
            Call ADODC_RESET
            
            Exit Sub
        End If
    End If
'
    Call BUTTON_ENABLE_SET(False)
'
    GRet = COPYFOLDER(wsRet, wGDir)
    If GRet <> True Then
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
'        Call BUTTON_ENABLE_SET(True)
        MsgBox GProduct & "フォルダの保存作業ができませんでした", vbExclamation
    
        Exit Sub
    End If
'
    '
    ' =========================================
    '                   後処理
    ' =========================================
    '----------< RESET Adodc >------------------------------------------------------
    Call ADODC_RESET
'    Call BUTTON_ENABLE_SET(True)
    MsgBox GProduct & "フォルダの保存作業か完了しました", vbInformation
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
システムバックアップ_ERR:
    pERR_MES = pPROGRAM_ID + "/ システムバックアップ() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' COPYFOLDER
'------------------------------------------------
Private Function COPYFOLDER(pPath As String, pDir As String)
'
    Dim Fso As FileSystemObject
'
    COPYFOLDER = False
'
    '金剛石 or 借換たろう！
    If GProduct <> "金剛石" Then
        pPath = pPath & "\" & Format(Date, "yyyymmdd") & "backup借換たろう"
    Else
        pPath = pPath & "\" & Format(Date, "yyyymmdd") & "backup金剛石"
    End If
'
    On Error GoTo Err_Hundle
'
    Set Fso = New FileSystemObject
    ' ディレクトリが存在しているかどうか判断する
    If Fso.FolderExists(pPath) = True Then
        GRet = MsgBox("フォルダが存在します。上書きしますか？", vbYesNo + vbExclamation, GProduct & "フォルダの保存")
        If GRet <> vbYes Then
            '----------< RESET Adodc >----------------------------------------------
            Call ADODC_RESET
            MsgBox GProduct & "フォルダを保存できませんでした", vbExclamation
    
            Exit Function
        End If
    End If

    ' ディレクトリをコピーする
    Call Fso.COPYFOLDER(pDir, pPath, True)

    DoEvents
    
    Set Fso = Nothing
    
    On Error GoTo 0
'
    COPYFOLDER = True
'
Exit Function
'----------< ERROR ROUTINE >--------------------------------------------------------
Err_Hundle:
    Resume Err_Hundle_END
Err_Hundle_END:
    Exit Function
End Function

'------------------------------------------------
' 閉じる_Click
'------------------------------------------------
Private Sub 閉じる_Click()
'
    '----------< DELETE_NEWREC >----------------------------------------------------
    If wNew企業名Key <> "" Then
        Call DELETE_NEWREC
    End If
'
    DoEvents
'
    Call FLG_KEYEND
'
    DoEvents
'
    GDb2.Close
    Set GDb2 = Nothing
    
    '----------< LOG WRITE >--------------------------------------------------------
    GStr = "4," & "9," & P8.FCChr(GMyComputerName, 16) & "," & P8.FCChr("", 30) & "," & P8.FCChr("", 30)
    GRet = PUT_LOG_FILE(GStr)
'
    End
'
End Sub

''------------------------------------------------
'' ログ一覧表
''------------------------------------------------
'Private Sub ログ一覧表()
''
'    '----------< DELETE_NEWREC >----------------------------------------------------
'    If wNew企業名Key <> "" Then
'        Call DELETE_NEWREC
'    End If
''
'    '----------< FLG 処理 >---------------------------------------------------------
'    GRet = FLG_CHECK_NYURYOKUOFF(企業名Key)
''
'    If B_6.Caption <> "ログ一覧表" Then
'        '----------< RESET Adodc >--------------------------------------------------
'        Call ADODC_RESET
'        Exit Sub
'    End If
''
'    GRpt.帳票名 = "ログ一覧表"
'    FBC010_ログ一覧表.Show vbModal
''
'    '----------< RESET Adodc >------------------------------------------------------
'    Call ADODC_RESET
''
'End Sub

'------------------------------------------------
' 完全削除
'------------------------------------------------
Private Sub 完全削除()
'
    Dim wsDB名 As String
    Dim ws01 As String
'
    On Error GoTo 完全削除_ERR
'
    ws01 = "他端末で実行している" & GProduct & "を終了してください"
    GRet = MsgBox(ws01 + vbCrLf + vbCrLf + "削除フラグの入っている企業データを物理的に消去します" + vbCrLf + "よろしいですか？", _
                         vbExclamation + vbOKCancel, "完全削除")
    If GRet = vbCancel Then
        '----------< RESET Adodc >--------------------------------------------------
        Call ADODC_RESET
        
        Exit Sub
    End If
'
    '----------< BUTTON_ENABLE_SET >------------------------------------------------
    Call BUTTON_ENABLE_SET(False)
'
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAA070_企業名マスタ "
    wstr = wstr + " Where not(削除日 is null)"
    Call AdoRecordsetOpen(GDb2, wRs2, wstr)
    If wRs2.eof Then
        MsgBox ("削除対象がありません")
        wRs2.Close
        Set wRs2 = Nothing
        
        ' =========================================
        '                   後処理
        ' =========================================
        Call ADODC_RESET
'        Call BUTTON_ENABLE_SET(True)
        
        Exit Sub
    Else
        Do While wRs2.eof = False
            wsDB名 = P8.FCStr(wRs2("DB名"))
            If Dir(GSerDir + "\" + wsDB名) <> "" _
                And (P8.FCStr(wRs2("入力中端末名")) = "" And P8.FCStr(wRs2("処理中端末名")) = "") Then
                
                '----------< AdoDbOpen_Check >--------------------------------------
                GRet = ADODBOPEN_CHECK("Jet", GDb, GSerDir + "\" + wsDB名, "", , GPwd, "排他")
                If GRet = True Then
                    '----------< DeleteFile >----------------------------------------
                    GRet = DeleteFile(GSerDir + "\" + wsDB名)
                    '----------< Delete List data >----------------------------------
                    wRs2.Delete
                End If
                DoEvents
            End If
            wRs2.MoveNext
        Loop
        wRs2.Close
        Set wRs2 = Nothing
'
        '----------< LOG WRITE >--------------------------------------------------
        GStr = "9," & "3," & P8.FCChr(GMyComputerName, 16) & "," & P8.FCChr("", 30) & "," & P8.FCChr("", 30)
        GRet = PUT_LOG_FILE(GStr)
'
        ' =========================================
        '                   後処理
        ' =========================================
        Call ADODC_RESET
'        Call BUTTON_ENABLE_SET(True)
        MsgBox "削除済みデータの物理削除が完了しました", vbInformation
    End If
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
完全削除_ERR:
    pERR_MES = pPROGRAM_ID + "/ 完全削除() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' COMDLG
'------------------------------------------------
Private Function COMDLG(wTitle As String, wDir As String, wFilter As String, wFile) As String
'
On Error GoTo ComCancel
    CommonDialog1.DialogTitle = wTitle
    CommonDialog1.InitDir = wDir
    CommonDialog1.Filter = wFilter
    CommonDialog1.FileName = wFile
    CommonDialog1.CancelError = True
    
    CommonDialog1.ShowSave
    COMDLG = CommonDialog1.FileName

    Exit Function
ComCancel:
    COMDLG = ""
End Function

'------------------------------------------------
' BrowseFolder
'------------------------------------------------
Public Function BrowseFolder(pMSG As String) As String
'データのあるフォルダの設定
      Dim Browse As BROWSEINFO
      Dim pID As Long
      Dim PathName As String
      Dim wi01 As Integer
'
      With Browse
            .hWndOwner = Me.hWnd
            .pidlRoot = CSIDL_DESKTOP
            .lpszTitle = pMSG
            .ulFlags = BIF_RETURNONLYFSDIRS
      End With

      '「フォルダの参照」ダイアログの呼び出し
      pID = SHBrowseForFolder(Browse)

      If pID Then
      '予めNull文字をセット
            PathName = String$(pMAX_PATH, vbNullChar)

      'SHBrowseForFolderで得られた値からフォルダのパス名を取得
            SHGetPathFromIDList pID, PathName

      '割り当てられたメモリを開放
            CoTaskMemFree pID

            wi01 = InStr(PathName, vbNullChar)
            If wi01 Then
                BrowseFolder = Left$(PathName, wi01 - 1)
            End If
      End If
'
End Function

'------------------------------------------------
' FileMaker
'------------------------------------------------
Private Function FileMaker(F_FileName As String, T_FileName As String) As Boolean
'
    Dim wJET As New JetEngine
'
    FileMaker = False
'
    On Error GoTo Err_Hundle
        wJET.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0" & _
                                ";Data Source=" & F_FileName & _
                                ";Persist Security Info=False" & _
                                ";Jet OLEDB:Database Password=" & GPwd, _
                            "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                ";Data Source=" & T_FileName & _
                                ";Jet OLEDB:Database Password=" & GPwd
'
        '----------< ファイル属性設定 >---------------------------------------------
        SetAttr T_FileName, vbNormal
    On Error GoTo 0
'
    FileMaker = True
'
Exit Function
'----------< ERROR ROUTINE >--------------------------------------------------------
Err_Hundle:
    Resume Err_Hundle_END
Err_Hundle_END:
    Exit Function
End Function

'------------------------------------------------
' 禁則文字
'------------------------------------------------
Private Function 禁則文字(FName As String) As Boolean
'
    禁則文字 = False
'
    If InStr(1, FName, "\") <> 0 Then
        Exit Function
    End If
    If InStr(1, FName, ":") <> 0 Then
        Exit Function
    End If
    If InStr(1, FName, ",") <> 0 Then
        Exit Function
    End If
    If InStr(1, FName, ";") <> 0 Then
        Exit Function
    End If
    If InStr(1, FName, "*") <> 0 Then
        Exit Function
    End If
    If InStr(1, FName, "?") <> 0 Then
        Exit Function
    End If
    If InStr(1, FName, """") <> 0 Then
        Exit Function
    End If
    If InStr(1, FName, "<") <> 0 Then
        Exit Function
    End If
    If InStr(1, FName, ">") <> 0 Then
        Exit Function
    End If
    If InStr(1, FName, "|") <> 0 Then
        Exit Function
    End If
'
    禁則文字 = True
'
End Function

'------------------------------------------------
' KIGYOSHORI_END
'------------------------------------------------
Private Sub KIGYOSHORI_END(pKeyName As String, pSdate As String)
'
    '----------< DAAA070_企業名マスタ >--------------------------------------------
    wstr = "Update "
    wstr = wstr + "DAAA070_企業名マスタ"
    wstr = wstr + " Set"
    wstr = wstr + " 入力中端末名 = '',"
    wstr = wstr + " 処理中端末名 = '',"
    wstr = wstr + " 端末コンピュータ名 = '" + GMyComputerName + "',"
    wstr = wstr + " 処理終了日付 = #" + pSdate + "#"
    wstr = wstr + " Where 企業名Key = '" & pKeyName & "'"
    GDb2.Execute wstr
'
    '----------< DAAA020_稼動中 >---------------------------------------------------
    wstr = "Update "
    wstr = wstr + "DAAA020_稼動中"
    wstr = wstr + " Set"
    wstr = wstr + " 稼動中フラグ = 0,"
    wstr = wstr + " 処理終了日付 = #" + pSdate + "#"
    wstr = wstr + " Where 端末コンピュータ名 = '" & GMyComputerName & "'"
    GDb2.Execute wstr
'
End Sub

'------------------------------------------------
' FLG_OFFMYPC
'------------------------------------------------
Private Sub FLG_OFFMYPC()
'
    '----------< DAAA070_企業名マスタ >--------------------------------------------
    wstr = "Update"
    wstr = wstr + " DAAA070_企業名マスタ"
    wstr = wstr + " Set"
    wstr = wstr + " 入力中端末名 = ''"
    wstr = wstr + " Where 入力中端末名 = '" & GMyComputerName & "'"
    GDb2.Execute wstr
    
    wstr = "Update"
    wstr = wstr + " DAAA070_企業名マスタ"
    wstr = wstr + " Set"
    wstr = wstr + " 処理中端末名 = ''"
    wstr = wstr + " Where 処理中端末名 = '" & GMyComputerName & "'"
    GDb2.Execute wstr
'
    '----------< DAAA020_稼動中 >---------------------------------------------------
    wstr = "Update"
    wstr = wstr + " DAAA020_稼動中"
    wstr = wstr + " Set"
    wstr = wstr + " 稼動中フラグ = 0"
    wstr = wstr + " Where 端末コンピュータ名 = '" & GMyComputerName & "'"
    wstr = wstr + " And 稼動中フラグ <> 0"
    GDb2.Execute wstr
'
End Sub

'------------------------------------------------
' FLG_CHECK_NYURYOKUOFF
'------------------------------------------------
Private Function FLG_CHECK_NYURYOKUOFF(pKeyName As String)
'
    Dim ws01 As String
'
    On Error GoTo FLG_CHECK_NYURYOKUOFF_ERR
'
    FLG_CHECK_NYURYOKUOFF = False
'
    If pKeyName = "" Then
        FLG_CHECK_NYURYOKUOFF = True
        Exit Function
    End If
'
FLG_CHECK_NYURYOKUOFF_ERR_RETRY:
'
    '----------< DAAA070_企業名マスタ >--------------------------------------------
    wstr = "Update"
    wstr = wstr + " DAAA070_企業名マスタ"
    wstr = wstr + " Set"
    wstr = wstr + " 入力中端末名 = '',"
    wstr = wstr + " 処理中端末名 = ''"
    wstr = wstr + " Where 企業名Key = '" & pKeyName & "'"
    wstr = wstr + " And (入力中端末名 = '' Or 入力中端末名 = '" & GMyComputerName & "')"
    wstr = wstr + " And (処理中端末名 = '' Or 処理中端末名 = '" & GMyComputerName & "')"
    GDb2.Execute wstr
'
On Error GoTo 0
'
    FLG_CHECK_NYURYOKUOFF = True
'
    Exit Function
'
'----------< ERROR ROUTINE >--------------------------------------------------------
FLG_CHECK_NYURYOKUOFF_ERR:
    If Err.Number = -214727887 Or Err.Number = -2147217887 Then
        Sleep (1000)
        Resume FLG_CHECK_NYURYOKUOFF_ERR_RETRY
    End If
    '
    pERR_MES = pPROGRAM_ID + "/ FLG_CHECK_NYURYOKUOFF() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                "他端末で同一企業(" & pKeyName & ")使用中です"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)
    '
    '----------< ADODC_RESET >------------------------------------------------------
    Call ADODC_RESET
    '
    Resume FLG_CHECK_NYURYOKUOFF_ERR_END
FLG_CHECK_NYURYOKUOFF_ERR_END:
    Exit Function
'
End Function

'------------------------------------------------
' FLG_NYURYOKUOFF
'------------------------------------------------
Private Sub FLG_NYURYOKUOFF(pKeyName As String)
'
    '----------< DAAA070_企業名マスタ >--------------------------------------------
    wstr = "Update"
    wstr = wstr + " DAAA070_企業名マスタ"
    wstr = wstr + " Set"
    wstr = wstr + " 入力中端末名 = ''"
    wstr = wstr + " Where 企業名Key = '" & pKeyName & "'"
    GDb2.Execute wstr
'
End Sub

'------------------------------------------------
' FLG_KIGYOSHORI
'------------------------------------------------
Private Function FLG_KIGYOSHORI(pKeyName As String, pSdate As String) As Boolean
'
    Dim wi01 As Integer
    Dim wSDate As String
    Dim ws01 As String, ws02 As String
'
    On Error GoTo FLG_KIGYOSHORI_ERR
'
    FLG_KIGYOSHORI = False
'
FLG_KIGYOSHORI_ERR_RETRY:

    wi_Kosin1 = 0: wi_Kosin2 = 0
'
    '----------< DAAA070_企業名マスタ >---------------------------------------------
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAA070_企業名マスタ"
    wstr = wstr + " Where 企業名Key = '" & pKeyName & "'"
    Call AdoRecordsetOpen(GDb2, wRs2, wstr)
    If wRs2.eof Then
        wRs2.AddNew
            
        wRs2("企業名Key") = pKeyName
        wRs2("最新処理日") = P8.FCDate(pSdate)
        wRs2("処理中端末名") = GMyComputerName
        wRs2("処理開始日付") = P8.FCDate(pSdate)
        wRs2("処理終了日付") = P8.FCDate(pSdate)
        wRs2("端末コンピュータ名") = GMyComputerName
            
        wRs2.Update
    Else
        ws01 = P8.FCStr(wRs2("入力中端末名"))
        If ws01 <> "" And ws01 <> GMyComputerName Then
            '----------< Msg >--------------------------------------------------
            If ws01 = "" Then
                ws02 = "他端末で同一企業(" & pKeyName & ")入力中です"
            Else
                ws02 = "他端末(" & ws01 & ")で同一企業(" & pKeyName & ")入力中です"
            End If
            GRet = MsgBox(ws02 & vbCrLf & Msg_01, vbOKOnly + vbExclamation)
            wRs2.Close
            Set wRs2 = Nothing
                
            Exit Function
        End If
    
        ws01 = P8.FCStr(wRs2("処理中端末名"))
        If ws01 <> "" And ws01 <> GMyComputerName Then
            wRs2.Close
            Set wRs2 = Nothing
            
            '----------< Msg >--------------------------------------------------
            If ws01 = "" Then
                ws02 = "他端末で同一企業(" & pKeyName & ")入力中です"
            Else
                ws02 = "他端末(" & ws01 & ")で同一企業(" & pKeyName & ")入力中です"
            End If
            GRet = MsgBox(ws02 & vbCrLf & Msg_01, vbOKOnly + vbExclamation)
                
            Exit Function
        End If
    
        wi01 = P8.FCDbl(wRs2("更新回数"))
        wi_Kosin1 = wi01
        w企業名Key = pKeyName
        
        wRs2("最新処理日") = P8.FCDate(pSdate)
        wRs2("入力中端末名") = ""
        wRs2("処理中端末名") = GMyComputerName
        wRs2("処理開始日付") = P8.FCDate(pSdate)
        wRs2("処理終了日付") = P8.FCDate(pSdate)
        wRs2("端末コンピュータ名") = GMyComputerName
            
        wRs2.Update
    End If
    wRs2.Close
    Set wRs2 = Nothing
        
    '----------< DAAA020_稼動中 >---------------------------------------------------
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAA020_稼動中"
    wstr = wstr + " Where 端末コンピュータ名 = '" & GMyComputerName & "'"
    Call AdoRecordsetOpen(GDb2, wRs2, wstr)
    If wRs2.eof Then
        wRs2.AddNew
        wRs2("端末コンピュータ名") = GMyComputerName
    End If
    
        wRs2("稼動中フラグ") = 1
        wRs2("処理開始日付") = P8.FCDate(pSdate)
        wRs2("処理終了日付") = P8.FCDate(pSdate)
        
        wRs2.Update
    wRs2.Close
    Set wRs2 = Nothing

On Error GoTo 0
'
    FLG_KIGYOSHORI = True
'
    Exit Function
'
'----------< ERROR ROUTINE >--------------------------------------------------------
FLG_KIGYOSHORI_ERR:
    If Err.Number = -214727887 Or Err.Number = -2147217887 Then
        Sleep (1000)
        Resume FLG_KIGYOSHORI_ERR_RETRY
    End If
    '
    pERR_MES = pPROGRAM_ID + "/ FLG_KIGYOSHORI() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                "他端末で同一企業(" & pKeyName & ")使用中です"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)
    '
    '----------< ADODC_RESET >------------------------------------------------------
    Call ADODC_RESET
    '
    Resume FLG_KIGYOSHORI_ERR_END
FLG_KIGYOSHORI_ERR_END:
    Exit Function
'
End Function

'------------------------------------------------
' FLG_TOROKUSET
'------------------------------------------------
Private Function FLG_TOROKUSET(pKeyName As String, Optional pNewrecord As String = "") As Boolean
'
    Dim wi01 As Integer, wi02 As Integer
    Dim wSDate As String
    Dim ws01 As String, ws02 As String
'
    On Error GoTo FLG_TOROKUSET_ERR
'
    FLG_TOROKUSET = False
'
    If pKeyName = "" Then
        FLG_TOROKUSET = True
        Exit Function
    End If
'
    FLG_New = False
    wi_Kosin1 = 0: wi_Kosin2 = 0
    wi01 = 0: wi02 = 0
    wSDate = Format(Now, "yyyy/mm/dd hh:mm:ss")
'
FLG_TOROKUSET_ERR_RETRY:
    '----------< DAAA070_企業名マスタ >---------------------------------------------
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAA070_企業名マスタ"
    wstr = wstr + " Where 企業名Key = '" & pKeyName & "'"
    Call AdoRecordsetOpen(GDb2, wRs2, wstr)
    If wRs2.eof Then
        If pNewrecord <> "" Then
            wRs2.AddNew
            
            wRs2("企業名Key") = pKeyName
            wRs2("企業名") = pKeyName
            wRs2("入力中端末名") = GMyComputerName
            wRs2("DB名") = pKeyName & ".mdb"
            wRs2("最新処理日") = P8.FCDate(wSDate)
            wRs2("作成日") = P8.FCDate(wSDate)
            wRs2("入力中端末名") = GMyComputerName
            wRs2("支店コード") = ""
            wRs2("支店名") = "単独企業"
            wRs2("親会社名") = ""
            wRs2("企業区分") = "単独企業"
            
            wRs2.Update
            
            FLG_New = True
            wNew企業名Key = pKeyName
            wi_Kosin1 = 0
        End If
    Else
        ws01 = P8.FCStr(wRs2("入力中端末名"))
        If ws01 <> "" And ws01 <> GMyComputerName Then
            '----------< Msg >------------------------------------------------------
            If ws01 = "" Then
                ws02 = "他端末で同一企業(" & pKeyName & ")入力中です"
            Else
                ws02 = "他端末(" & ws01 & ")で同一企業(" & pKeyName & ")入力中です"
            End If
            GRet = MsgBox(ws02 & vbCrLf & Msg_01, vbOKOnly + vbExclamation)
            wRs2.Close
            Set wRs2 = Nothing
                
            '----------< ADODC_RESET >----------------------------------------------
            Call ADODC_RESET
                
            Exit Function
        End If
    
        ws01 = P8.FCStr(wRs2("処理中端末名"))
        If ws01 <> "" And ws01 <> GMyComputerName Then
            wRs2.Close
            Set wRs2 = Nothing
            
            '----------< Msg >------------------------------------------------------
            If ws01 = "" Then
                ws02 = "他端末で同一企業(" & pKeyName & ")入力中です"
            Else
                ws02 = "他端末(" & ws01 & ")で同一企業(" & pKeyName & ")入力中です"
            End If
            GRet = MsgBox(ws02 & vbCrLf & Msg_01, vbOKOnly + vbExclamation)
                
            '----------< ADODC_RESET >----------------------------------------------
            Call ADODC_RESET
                
            Exit Function
        End If
    
        wi01 = P8.FCDbl(wRs2("更新回数"))
        wi_Kosin1 = wi01
        w企業名Key = pKeyName
        
        wRs2("入力中端末名") = GMyComputerName
        wRs2("処理中端末名") = ""
            
        wRs2.Update
    End If
    wRs2.Close
    Set wRs2 = Nothing
    
On Error GoTo 0
'
    FLG_TOROKUSET = True
'
    Exit Function
'
'----------< ERROR ROUTINE >--------------------------------------------------------
FLG_TOROKUSET_ERR:
    If Err.Number = -214727887 Or Err.Number = -2147217887 Then
        Sleep (1000)
        Resume FLG_TOROKUSET_ERR_RETRY
    End If
    '
    pERR_MES = pPROGRAM_ID + "/ FLG_TOROKUSET() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                "他端末で同一企業(" & pKeyName & ")使用中です"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)
    '
    '----------< ADODC_RESET >------------------------------------------------------
    Call ADODC_RESET
    '
    Resume FLG_TOROKUSET_ERR_END
FLG_TOROKUSET_ERR_END:
    Exit Function
'
End Function

'------------------------------------------------
' FLG_KEYEND
'------------------------------------------------
Private Sub FLG_KEYEND()
'
    On Error GoTo FLG_KEYEND_ERR
'
    '----------< DAAA070_企業名マスタ >--------------------------------------------
    wstr = "Update"
    wstr = wstr + " DAAA070_企業名マスタ"
    wstr = wstr + " Set"
    wstr = wstr + " 入力中端末名 = '',"
    wstr = wstr + " 処理中端末名 = ''"
    wstr = wstr + " Where 企業名Key = '" & 企業名Key & "'"
    GDb2.Execute wstr
'
    DoEvents
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
FLG_KEYEND_ERR:
    pERR_MES = pPROGRAM_ID + "/ FLG_KEYEND() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'------------------------------------------------
' ADODBOPEN_CHECK
'------------------------------------------------
Private Function ADODBOPEN_CHECK _
                   (pProvider As String, _
                    pAdoDb As ADODB.Connection, _
                    pDbName As String, _
                    Optional pSource As String = "", _
                    Optional pUID As String = "", _
                    Optional pPassword As String = "", _
                    Optional pMode As String = "") As Boolean
'
    Dim wstr As String
'
    ADODBOPEN_CHECK = False
'
    On Error GoTo Err_Hundle
        Select Case LCase(pProvider)
        Case "jet"
            wstr = "Provider=Microsoft.Jet.OLEDB.4.0"
            wstr = wstr & ";Data Source=" & pDbName
            wstr = wstr & ";Persist Security Info=False"
            wstr = wstr & ";Jet OLEDB:Database Password=" & pPassword
        
            pAdoDb.ConnectionString = wstr
    
            If pMode = "排他" Then
                pAdoDb.Mode = adModeShareExclusive
            Else
                pAdoDb.Mode = adModeUnknown
            End If
            
        End Select

        pAdoDb.Open
        
        pAdoDb.Close
        Set pAdoDb = Nothing
    On Error GoTo 0
'
    ADODBOPEN_CHECK = True
'
Exit Function
'----------< ERROR ROUTINE >--------------------------------------------------------
Err_Hundle:
    Resume Err_Hundle_END
Err_Hundle_END:
    Exit Function
End Function

'------------------------------------------------
' DELETE_NEWREC
'------------------------------------------------
Private Sub DELETE_NEWREC()
'
    On Error GoTo DELETE_NEWREC_ERR
'
    wstr = ""
    wstr = wstr + "Delete"
    wstr = wstr + " From DAAA070_企業名マスタ"
    wstr = wstr + " Where 企業名Key = '" + wNew企業名Key + "'"
    GDb2.Execute wstr
            
    wNew企業名Key = ""
    
    DoEvents
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
DELETE_NEWREC_ERR:
    pERR_MES = pPROGRAM_ID + "/ DELETE_NEWREC() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub


