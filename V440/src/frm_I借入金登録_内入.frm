VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_I借入金登録_内入 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "借入金登録 内入入力"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12870
   Icon            =   "frm_I借入金登録_内入.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "登録"
      Height          =   3255
      Left            =   7800
      TabIndex        =   5
      Top             =   720
      Width           =   4935
      Begin VB.CommandButton 削除 
         Caption         =   "削除"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   34
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton 閉じる 
         Caption         =   "閉じる(F12)"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton 登録 
         Caption         =   "登録（F11)"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   32
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox 年月日 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   2  'ｵﾌ
         Left            =   1800
         TabIndex        =   0
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox T_2 
         Alignment       =   1  '右揃え
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   2
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox T_1 
         Alignment       =   1  '右揃え
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   1
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox T_3 
         Alignment       =   1  '右揃え
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   3
         Top             =   1680
         Visible         =   0   'False
         Width           =   2655
      End
      Begin 借換たろう.ZU070_Label 新規変更 
         Height          =   375
         Left            =   120
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor_Shape1=   8454016
         BackColor_Shape2=   8421504
         BorderColor_Shape1=   49152
         BorderColor_Shape2=   4210752
         ForeColor       =   255
         Caption         =   "新規変更"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label12 
         Alignment       =   2  '中央揃え
         Caption         =   "円"
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
         Left            =   4440
         TabIndex        =   31
         Top             =   2160
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label11 
         Alignment       =   2  '中央揃え
         Caption         =   "円"
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
         Left            =   4440
         TabIndex        =   30
         Top             =   1800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label L1_4 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
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
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label L1_3 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
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
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label L1_年月日 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 年月日"
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
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   2  '中央揃え
         Caption         =   "円"
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
         Left            =   4440
         TabIndex        =   26
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   2  '中央揃え
         Caption         =   "円"
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
         Left            =   4440
         TabIndex        =   25
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label L1_2 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
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
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label L1_1 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
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
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label L_4 
         Alignment       =   1  '右揃え
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BorderStyle     =   1  '実線
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
         Left            =   1800
         TabIndex        =   22
         Top             =   2040
         Visible         =   0   'False
         Width           =   2685
      End
      Begin VB.Label L_3 
         Alignment       =   1  '右揃え
         BorderStyle     =   1  '実線
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
         Left            =   1800
         TabIndex        =   21
         Top             =   1680
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label L_1 
         Alignment       =   1  '右揃え
         BorderStyle     =   1  '実線
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
         Left            =   1800
         TabIndex        =   20
         Top             =   960
         Width           =   2655
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   8685
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   15319
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   0
      Top             =   9240
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
   Begin 借換たろう.ZU050_Button ZU050_Button1 
      Height          =   615
      Left            =   120
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "内入入力"
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
   Begin VB.Label L_合計融資残高 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
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
      Left            =   4320
      TabIndex        =   14
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label L_融資金額 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
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
      Left            =   4320
      TabIndex        =   15
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label L_最終返済年月 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
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
      Left            =   4320
      TabIndex        =   11
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label L_初回返済年月 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
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
      Left            =   4320
      TabIndex        =   10
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label L_番号 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
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
      Left            =   4320
      TabIndex        =   16
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label L_実行日 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
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
      Left            =   4320
      TabIndex        =   9
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label L_番号1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   " 借入番号"
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
      Left            =   2520
      TabIndex        =   19
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "円"
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
      TabIndex        =   18
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   " 融資金額"
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
      Left            =   2520
      TabIndex        =   17
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label L_合計融資残高1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   " 融資残高"
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
      Left            =   2520
      TabIndex        =   13
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label L_合計融資残高2 
      Alignment       =   2  '中央揃え
      Caption         =   "円"
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
      TabIndex        =   12
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label17 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   " 最終返済年月日"
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
      Left            =   2520
      TabIndex        =   8
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label18 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   " 初回返済年月日"
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
      Left            =   2520
      TabIndex        =   7
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label19 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   " 実行日"
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
      Left            =   2520
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "frm_I借入金登録_内入"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "frm_I借入金登録_内入"

Dim wRs As ADODB.Recordset
Dim wstr As String

Dim wslog As String

Dim w借入データ As MAA910_借入金
Dim w借入内入 As MAA910_借入金内入

Dim wdYkin As Double
Dim wFname As String, wFname2 As String
Dim wsTbl As String, wsTbl2 As String, wsTbl3 As String
Dim w初回 As Variant, w最終 As Variant
Dim wsBango As String

Dim wv実行日 As Variant, wv初回返済実行日 As Variant, wv最終返済実行日 As Variant
Dim wv解約実行日 As Variant
Dim wd前月残高 As Double, wd当月残高 As Double
Dim wv最新日 As Variant
Dim FLG_MAX As Boolean

Dim wi据置X回目 As Integer
'
'------------------------------------------------
' Form_Load
'------------------------------------------------
Private Sub Form_Load()
'
'    Me.Caption = GFcap
    Me.Left = G_LEFT
    Me.Top = G_TOP
'
    wFname = GStr
    wFname2 = GStr_3
    'ZU050_Button1.Caption = wFname & wFname2 & Space(1) & "登録"
    
    wsBango = GStr_2
    
    L_番号.Caption = wsBango
    
    登録.Caption = "金額設定" & vbCr & "登録/(F11)"
        
    Select Case wFname
    Case "借入金登録"
        L_番号1.Caption = " 借入番号"
        
        wsTbl = "DBDA010_借入金明細TR"
        wsTbl2 = "DBDA010_借入金"
        
    Case "貸付登録"
        L_番号1.Caption = " 貸付番号"
    
        wsTbl = "DBDA010_貸付金明細TR"
        wsTbl2 = "DBDA010_貸付金"
    
    End Select
    
    Select Case wFname2
    Case "内入入力"
        wsTbl3 = "DBDA010_借入金内入"
    End Select
    
    GStr = "": GStr_1 = "": GStr_2 = ""
    GStr_3 = ""
'
    ' =========================================
    '                 初期設定
    ' =========================================
'
    wv実行日 = Null
    wv初回返済実行日 = Null
    wv最終返済実行日 = Null
    wv解約実行日 = Null

    L_実行日.Caption = ""
    L_初回返済年月.Caption = ""
    L_最終返済年月.Caption = ""
    L_融資金額.Caption = ""
    wdYkin = 0
    '取消 = 0
    
    wstr = ""
    wstr = wstr + "Select"
    wstr = wstr + " 借入番号,実行日,初回返済年月,最終返済年月,融資金額,"
    wstr = wstr + " 初回返済実行日,最終返済実行日,解約実行日,"
    wstr = wstr + " 利息区分"
    wstr = wstr + " From " & wsTbl2
    wstr = wstr + " Where 借入番号 = '" & wsBango & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
        L_実行日.Caption = Format(P8.FCStr(wRs("実行日")), Gfmt年月日)
        L_初回返済年月.Caption = Format(P8.FCStr(wRs("初回返済年月")), Gfmt年月)
        L_最終返済年月.Caption = Format(P8.FCStr(wRs("最終返済年月")), Gfmt年月)
        wdYkin = P8.FCDbl(wRs("融資金額"))
        L_融資金額.Caption = Format(wdYkin, "#,##0")
                        
        If P8.FCDbl(wRs("利息区分")) = XMXA020_区分("利息区分", "利息先払") Then
            wi据置X回目 = 3
        Else
            wi据置X回目 = 1
        End If
        
        wv実行日 = wRs("実行日")
        wv初回返済実行日 = wRs("初回返済実行日")
        wv最終返済実行日 = wRs("最終返済実行日")
        wv解約実行日 = wRs("解約実行日")
    
    End If
    wRs.Close
    Set wRs = Nothing
'
    '合計融資残高
    L_合計融資残高.Caption = ""
    
    wstr = ""
    wstr = wstr + "Select 融資残高"
    wstr = wstr + " From " & wsTbl
    wstr = wstr + " Where 借入番号 = '" & wsBango & "'"
    wstr = wstr + " And 取消フラグ=0"
    wstr = wstr + " Order by 実際年月日 desc"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
        L_合計融資残高.Caption = Format(P8.FCDbl(wRs("融資残高")), "#,##0")
    End If
    
    wRs.Close
    Set wRs = Nothing
'
    '内入
    FLG_MAX = False
    
    L_合計融資残高.Caption = ""
    L_合計融資残高1.Caption = " 融資残高"
    L_合計融資残高2.Caption = "円"
    
    L_合計融資残高.Visible = True
    L_合計融資残高1.Visible = True
    L_合計融資残高2.Visible = True

    '
    L1_年月日.Caption = " 返済年月日"
    L1_1.Caption = " 元金"
    L1_2.Caption = " 利息額"
    L1_3.Caption = " 返済金額"
    L1_4.Caption = " 融資残高"
    
    L_1.Caption = ""
    L_3.Caption = ""
    L_4.Caption = ""
    
    年月日.Text = ""
    T_1.Text = ""
    T_2.Text = ""
    T_3.Text = ""
    
    T_1.Visible = True
    T_3.Visible = False
    
    Select Case wFname2
    Case "内入入力"
        
        L_初回返済年月.Caption = Format(wv初回返済実行日, Gfmt年月日)
        L_最終返済年月.Caption = Format(wv最終返済実行日, Gfmt年月日)
        
        L1_年月日.Caption = " 内入年月日"
        L1_1.Caption = " 前月残高"
        L1_2.Caption = " 内入金額"
        L1_3.Caption = " 手数料"
        L1_4.Caption = " 当月残高"
        
        T_1.Visible = False
'        T_3.Visible = True
            
        L_合計融資残高.Caption = ""
        L_合計融資残高1.Caption = ""
        L_合計融資残高2.Caption = ""
        
        L_合計融資残高.Visible = False
        L_合計融資残高1.Visible = False
        L_合計融資残高2.Visible = False
    
        'ワークテーブル作成とワークデータセット
        Call 内入ワークテーブル作成
        
    End Select
'
    Call 登録後初期セット
'
End Sub

'------------------------------------------------
' Form_Activate
'------------------------------------------------
Private Sub Form_Activate()
'
    DoEvents

    'Call MXA010_検索用データクリア
    Call CEkey.AllSelect
End Sub

'------------------------------------------------
' Form_KeyDown
'------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
    If KeyCode = vbKeyF11 Then
        Call 登録_Click
    End If
    
    If KeyCode = vbKeyF12 Then
        Call 閉じる_Click
    End If
'
End Sub

'------------------------------------------------
' Form_KeyPress
'------------------------------------------------
Private Sub Form_KeyPress(KeyAscii As Integer)
'
    KeyAscii = CEkey.X020_EnterKey(Me, KeyAscii, True)
'
End Sub

'------------------------------------------------
' AdodcRefresh
'------------------------------------------------
Private Sub AdodcRefresh()
'
    On Error GoTo AdodcRefresh_ERR
'
    ' =========================================
    '             グッリドの初期値
    ' =========================================
'    Call MXA030_DataGridInit(DataGrid1)
    DataGrid1.AllowRowSizing = False
    DataGrid1.HeadFont.Size = 9
    DataGrid1.HeadFont.Bold = True
    DataGrid1.Font.Size = 9
    DataGrid1.BackColor = C_Yellow
    DataGrid1.ForeColor = RGB(0, 0, 160)
    
    Set DataGrid1.DataSource = Adodc1
  
    ' =========================================
    '              ConnectionString
    ' =========================================
    Call AdodcSet(Adodc1, GDb)
  
    ' =========================================
    '              メインクエリ
    ' =========================================
    GWhere = ""
    GWhere = GWhere & " And 借入番号='" & wsBango & "'"
    
    GWhere = " Where (1=1) " + GWhere
    
    wstr = ""
    wstr = wstr + "Select"
    
    wstr = wstr + " IIF(取消フラグ２=0,1,2),"
    wstr = wstr + " IIF(取消フラグ=0,1,2),"
    
    wstr = wstr + " 借入番号,実際年月日,元金額,利息額,返済金額,融資残高,取消フラグ,"
    wstr = wstr + " Format(実際年月日,'" & Gfmt年月日 & "') As Grd年月日,"
    wstr = wstr + " Format(元金額,'#,##0') As Grd元金,"
    wstr = wstr + " Format(利息額,'#,##0') As Grd利息額,"
    wstr = wstr + " Format(返済金額,'#,##0') As Grd返済金額,"
    wstr = wstr + " Format(融資残高,'#,##0') As Grd融資残高,"
    wstr = wstr + " IIF(取消フラグ = 0,'','×') As Grd取消,"
    wstr = wstr + " IIF(取消フラグ２ = 0,'','×') As Grd取消2"
    wstr = wstr + " From " & wsTbl
    wstr = wstr + GWhere
    wstr = wstr + " Order By 1,2,実際年月日"
    
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        Call XZMA010_DataGrid_Set("年月日", "返済年月日", 1800, "L")
        Call XZMA010_DataGrid_Set("元金", "", 1500, "R")
        Call XZMA010_DataGrid_Set("利息額", "", 1500, "R")
        Call XZMA010_DataGrid_Set("返済金額", "", 1500, "R")
        Call XZMA010_DataGrid_Set("融資残高", "", 1500, "R")
        Call XZMA010_DataGrid_Set("取消", "", 550, "C")
        Call XZMA010_DataGrid_Set("取消2", "", 550, "C")
    Call XZMA010_DataGrid_Action(DataGrid1)
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
' 内入ワークテーブル作成
'------------------------------------------------
Private Sub 内入ワークテーブル作成()
'
    Dim wRs1 As ADODB.Recordset
    Dim wstr1 As String
    
    Dim wiCnt As Integer
    Dim j As Integer, k As Integer
    Dim ws01 As String
'
    On Error GoTo 内入ワークテーブル作成_ERR
'
    '----------< ワークテーブル削除 >------------------------------------------
    wstr = "Delete * From DCDA020_借入金明細"
    GDb.Execute wstr
    
    wstr = "Delete * From DCHA010_Gridワーク"
    GDb.Execute wstr
'
    If wsBango = "" Then
        Exit Sub
    End If
'
    '----------< テーブル Write >----------------------------------------------
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From " & wsTbl2 & " As k"
    wstr = wstr + " Where K.借入番号 = '" & wsBango + "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
        
            w借入データ = MBD010_借入データセット(wRs)
            Call MBD010_借入金テーブル作成("", w借入データ)
            Call MBD010_借入明細作成("", w借入データ)
  
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    wstr = ""
    wstr = wstr & "Insert Into DCHA010_Gridワーク"
    wstr = wstr & "(テキスト1,年月日1,数値1,数値2,数値3,数値4)"
    wstr = wstr & " Select "
    wstr = wstr & "M.借入番号,M.実際年月日,M.元金額,M.手数料,M.融資残高,据置X回目"
    wstr = wstr & " From DCDA020_借入金明細 As M"
    wstr = wstr & " INNER JOIN DBDA010_借入金 As K"
    wstr = wstr & " ON M.借入番号 = K.借入番号"
    
    wstr = wstr & " Where (M.借入番号='" & wsBango & "' And M.据置X回目 = 1)"
    wstr = wstr & " OR (M.借入番号='" & wsBango & "' AND M.据置X回目=3)"
    wstr = wstr & " OR (M.借入番号='" & wsBango & "' AND M.据置X回目=4)"
    wstr = wstr & " OR (M.借入番号='" & wsBango & "' AND M.据置X回目=2"
    wstr = wstr & " And (M.元金額<>0 or M.利息額<>0))"
    'wstr = wstr & " And (K.プロジェクト番号='' or K.プロジェクト番号 is null) AND (M.元金額<>0 or M.利息額<>0))"
    'wstr = wstr & " And K.プロジェクト番号='' AND M.元金額<>0)"
    
    GDb.Execute wstr
'
    wv最新日 = Null
    FLG_MAX = False
'
    '***借入金内入
    Call MBD010_借入内入_クリア
    
    w借入内入 = MBD010_借入内入(wsBango)
'
    j = 0
    
    For k = 1 To 500 '内入回数
        j = j + 1
        
        If IsNull(w借入内入.内入(k).内入x回目年月日) Or P8.FCStr(w借入内入.内入(k).内入x回目年月日) = "" Then
            Exit For
        End If
        
        If IsDate(w借入内入.内入(k).内入x回目年月日) Then
            wv最新日 = w借入内入.内入(k).内入x回目年月日
            wiCnt = j
        End If
    
    Next k
    
    If wiCnt >= 500 Then
        FLG_MAX = True
    Else
        FLG_MAX = False
    End If
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
内入ワークテーブル作成_ERR:
    pERR_MES = pPROGRAM_ID + "/ 内入ワークテーブル作成() でエラー" + vbCrLf + vbCrLf + _
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
' AdodcRefresh_Uchiire
'------------------------------------------------
Private Sub AdodcRefresh_Uchiire()
'
    On Error GoTo AdodcRefresh_Uchiire_ERR
'
    ' =========================================
    '             グッリドの初期値
    ' =========================================
'    Call MXA030_DataGridInit(DataGrid1)
    DataGrid1.AllowRowSizing = False
    DataGrid1.HeadFont.Size = 9
    DataGrid1.HeadFont.Bold = True
    DataGrid1.Font.Size = 9
    DataGrid1.BackColor = C_Yellow
    DataGrid1.ForeColor = RGB(0, 0, 160)
    
    Set DataGrid1.DataSource = Adodc1
  
    ' =========================================
    '              ConnectionString
    ' =========================================
    Call AdodcSet(Adodc1, GDb)
  
    ' =========================================
    '              メインクエリ
    ' =========================================
    GWhere = ""
    GWhere = " Where (1=1) " + GWhere
    
    wstr = ""
    wstr = wstr + "Select"
    wstr = wstr + " (SELECT Count(*) FROM DCHA010_Gridワーク AS T2"
    wstr = wstr + " WHERE format(DCHA010_Gridワーク.年月日1,'yyyy/mm/dd') > format(T2.年月日1,'yyyy/mm/dd'))+1 As GrdNo,"
    wstr = wstr + " format(年月日1,'" & Gfmt年月日 & "') As Grd年月日,"
    wstr = wstr + " IIF(数値4=2,'','○') As Grd内入,"
    wstr = wstr + " format(数値1+数値3,'#,##0') As Grd前月残高,"
    wstr = wstr + " format(数値1,'#,##0') As Grd内入金額,"
    wstr = wstr + " format(数値2,'#,##0') As Grd手数料,"
    wstr = wstr + " format(数値3,'#,##0') As Grd当月残高"
    wstr = wstr + " From DCHA010_Gridワーク"
    wstr = wstr + GWhere
    wstr = wstr + " Order By 年月日1"
    
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        Call XZMA010_DataGrid_Set("No", "回", 300, "C")
        Call XZMA010_DataGrid_Set("年月日", "返済日", 1400, "R")
        Call XZMA010_DataGrid_Set("内入", "", 500, "C")
        Call XZMA010_DataGrid_Set("前月残高", "", 1600, "R")
        Call XZMA010_DataGrid_Set("内入金額", "", 1600, "R")
'        Call XZMA010_DataGrid_Set("手数料", "", 1200, "R")
        Call XZMA010_DataGrid_Set("当月残高", "", 1600, "R")
    Call XZMA010_DataGrid_Action(DataGrid1)
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
AdodcRefresh_Uchiire_ERR:
    pERR_MES = pPROGRAM_ID + "/ AdodcRefresh_Uchiire() でエラー" + vbCrLf + vbCrLf + _
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
    Call CEkey.SetFs(年月日, True)
End Sub

'------------------------------------------------
' DataGrid1_LostFocus
'------------------------------------------------
Private Sub DataGrid1_LostFocus()
'
    Dim ws01 As String
'
    On Error Resume Next
        Dim wCheckValue As Variant
        wCheckValue = Adodc1.Recordset.Fields.Item("Grd年月日")
        If Err.Number = 3021 Then GoTo Exit_Sub
    On Error GoTo Err_Hundle
        年月日 = P8.FCStr(Adodc1.Recordset.Fields.Item("Grd年月日"))
    On Error GoTo 0
'
    Select Case wFname2
    Case "内入入力"
        Call 画面セット_Uchiire(True)
    Case Else
        Call 画面セット(True)
    End Select
'
    If DataGrid1.Splits.Count <> 1 Then
        DataGrid1.Splits.Remove 1
    End If

    Call CEkey.SetFs(年月日, True)

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
' B_SET_Click
'------------------------------------------------
Private Sub B_SET_Click()
'
    Select Case wFname2
    Case "内入入力"
        Call 画面セット_Uchiire(False)
    Case Else
        Call 画面セット(False)
    End Select
'
    Call CEkey.AllSelect
'
End Sub

'------------------------------------------------
' 画面セット
'------------------------------------------------
Private Function 画面セット(pGridClick As Boolean) As Boolean
'
    Dim j As Integer
    Dim ws01 As String
'
    On Error GoTo 画面セット_ERR
'
    画面セット = False
'
    ' =========================================
    '                画面クリア
    ' =========================================
    T_1 = ""
    T_2 = ""
    L_3.Caption = ""
    L_4.Caption = ""
    '取消 = 0
    
    ' =========================================
    '            借入金マスタ セット
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月", 年月日.Text)
    If GVar1 = 0 Then
        GVar1 = Null
    End If
    
    wstr = ""
    wstr = wstr + "Select"
    wstr = wstr + " 借入番号,実際年月日,元金額,利息額,返済金額,融資残高,"
    wstr = wstr + " 取消フラグ,取消フラグ２"
    wstr = wstr + " From " & wsTbl
    wstr = wstr + " Where 借入番号 = '" & wsBango & "'"
    wstr = wstr + " And Format(実際年月日,'yyyymmdd') = '" & Format(GVar1, "yyyymmdd") & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.eof Then
            If 年月日 <> "" Then
'                GRet = MsgBox("新規レコードを追加します。よろしいですか？", vbYesNo)
'                If GRet = vbNo Then
'                    '新規変更.Caption = ""
'                    wRs.Close
'                    Set wRs = Nothing
'
'                    Exit Function
'                End If
                
                新規変更.Caption = "新規登録"
                Call CEkey.SetFs(T_1, True)
    
            End If
        Else
            画面セット = True
            Call CEkey.SetFs(T_1, True)
            新規変更.Caption = "変更"
            
            T_1 = P8.FFormat(wRs("元金額"), "#,##0")
            T_2 = P8.FFormat(wRs("利息額"), "#,##0")
            L_3.Caption = P8.FFormat(wRs("返済金額"), "#,##0")
            L_4.Caption = P8.FFormat(wRs("融資残高"), "#,##0")
            '取消 = wRs("取消フラグ")
        End If
    wRs.Close
    Set wRs = Nothing
'
    '------------------------------------------
    '          ** グリッドコントロール **
    '------------------------------------------
    If Not pGridClick Then
        DoEvents
        Call AdodcRefresh
    End If

    DoEvents
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd年月日 = '" + 年月日 + "'")
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
' 画面セット_Uchiire
'------------------------------------------------
Private Function 画面セット_Uchiire(pGridClick As Boolean) As Boolean
'
    Dim FLG_New As Boolean
'
    On Error GoTo 画面セット_Uchiire_ERR
'
    画面セット_Uchiire = False
'
    L_1.Caption = ""
    T_2 = 0
    T_3 = 0
    L_4.Caption = ""
    '取消 = 0
    
    FLG_New = False

    wd前月残高 = 0
    wd当月残高 = 0
    
    ' =========================================
    '                画面クリア
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月", 年月日)
    If GVar1 = 0 Then
        GVar1 = Null
    End If
    
    wstr = ""
    wstr = wstr + "Select * From  DCHA010_Gridワーク"
    wstr = wstr + " Where Format(年月日1,'yyyymmdd') = '" & Format(GVar1, "yyyymmdd") & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.eof Then
        If Not IsNull(GVar1) Then
'            GRet = MsgBox("新規レコードを追加します。よろしいですか？", vbYesNo)
'            If GRet = vbNo Then
'                wRs.Close
'                Set wRs = Nothing
'
'                Exit Function
'            End If
            新規変更.Caption = "新規登録"
            If FLG_MAX = True Then
                GRet = MsgBox("内入回数500回を越えると登録できません。", vbOKOnly)
                
                wRs.Close
                Set wRs = Nothing
                
                Exit Function
            End If
            
            FLG_New = True
            Call CEkey.SetFs(T_2, True)
        End If
    Else
        新規変更.Caption = "変更"
        wd前月残高 = P8.FCDbl(wRs("数値1") + wRs("数値3"))
        wd当月残高 = P8.FCDbl(wRs("数値3"))
        
        年月日 = Format(wRs("年月日1"), Gfmt年月日)
        L_1.Caption = P8.FFormat(wd前月残高, "#,##0")
        T_2 = P8.FFormat(P8.FCDbl(wRs("数値1")), "#,##0")
        T_3 = P8.FFormat(P8.FCDbl(wRs("数値2")), "#,##0")
        L_4.Caption = P8.FFormat(wd当月残高, "#,##0")
        
    End If
    wRs.Close
    Set wRs = Nothing
'
    '前月残高セット
    If Not IsNull(GVar1) And FLG_New = True Then
        Call 内入残高_セット(CDate(Format(GVar1, "yyyy/mm/dd")))
    End If
'
    ' =========================================
    '            Grid セット
    ' =========================================
    If Not pGridClick Then
        DoEvents
        Call AdodcRefresh_Uchiire
    End If
    
    DoEvents
    
    If FLG_New = True Then
        Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd年月日 = '" + Format(wv最新日, Gfmt年月日) + "'")
    Else
        Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd年月日 = '" + 年月日 + "'")
    End If
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
画面セット_Uchiire_ERR:
    pERR_MES = pPROGRAM_ID + "/ 画面セット_Uchiire() でエラー" + vbCrLf + vbCrLf + _
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
' 登録後初期セット
'------------------------------------------------
Private Sub 登録後初期セット()
'
    Dim w年月日 As String
'
    年月日 = ""

    Select Case wFname2
    Case "内入入力"
        Call 画面セット_Uchiire(False)
    Case Else
        Call 画面セット(False)
    End Select
    
    '----------------------------------------
    '               更新行を表示
    '----------------------------------------
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd年月日 = '" + 年月日 + "'")
    Call CEkey.SetFs(T_1, True)
'
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
    GStr = wFname
    GStr_1 = wsBango
    
    Unload Me
    
    frm_I借入金登録.Enabled = True
    Call frm_I借入金登録.画面セット呼出
'
End Sub

Private Sub Form_Terminate()
    On Error Resume Next
    
    '----------< ワークテーブル削除 >------------------------------------------
    wstr = "Delete * From DCDA020_借入金明細"
    GDb.Execute wstr
    
    wstr = "Delete * From DCHA010_Gridワーク"
    GDb.Execute wstr
    
    On Error GoTo 0
End Sub

'------------------------------------------------
' 年月日_LostFocus
'------------------------------------------------
Private Sub 年月日_LostFocus()
'
    Dim ws01 As String
    Dim w年月日 As Date
'
    Call P8.FCControlLeft(年月日, 30)
    
    Select Case Screen.ActiveControl.Name
        Case "閉じる", "DataGrid1" ', "年月日"
            Exit Sub
    End Select
   
    If 年月日 = "" Then
'        MsgBox "コードを入力してください"
'        Call CEkey.SetFs(年月日, True)
        Exit Sub
    Else
        If InStrRev(年月日, "年") Then
            GVar1 = C年月日.平成To西暦("", 年月日)
            If GVar1 = 0 Then
                MsgBox "年月日を入力してください", vbExclamation
                年月日 = "": Call CEkey.SetFs(年月日, True)
                Exit Sub
            End If
        Else
'            If Len(年月日) < 5 Then
'                MsgBox "年月日を入力してください", vbExclamation
'                年月日 = "": Call CEkey.SetFs(年月日, True)
'                Exit Sub
'            End If
'
'            ws01 = Mid$(年月日 & "000000", 3, 2)
'            If ws01 < "01" Or ws01 > "12" Then
'                MsgBox "年月日を入力してください", vbExclamation
'                年月日 = "": Call CEkey.SetFs(年月日, True)
'                Exit Sub
'            End If
'
'            ws01 = Right$("000000" & 年月日, 2)
'            If ws01 < "01" Or ws01 > "31" Then
'                MsgBox "年月日を入力してください", vbExclamation
'                年月日 = "": Call CEkey.SetFs(年月日, True)
'                Exit Sub
'            End If

        End If
    End If
       
    年月日 = C年月日.FormatDate("年月日", 年月日)
    If C年月日.平成To西暦("年月", 年月日) = 0 Then
        MsgBox "年月日が違います", vbExclamation
        年月日 = "": Call CEkey.SetFs(年月日, True)
        Exit Sub
    End If
'
    Select Case wFname2
    Case "内入入力"
        GVar1 = C年月日.平成To西暦("年月", 年月日)
        If IsDate(GVar1) Then
            Call C休日.計算(CDate(GVar1), w借入データ.営業日区分)
            w年月日 = C休日.次回稼働日
            
            If Format(w年月日, "yyyy/mm/dd") <> Format(CDate(GVar1), "yyyy/mm/dd") Then
                
                GRet = MsgBox("内入年月日を稼働日(" & Format(w年月日, Gfmt年月日) & ")にセットします。よろしいですか？", vbYesNo)
                If GRet = vbYes Then
                    年月日 = Format(w年月日, Gfmt年月日)
                End If
            End If
        
            Call B_SET_Click
        End If
            
        Exit Sub
        
    End Select
'
    Select Case Screen.ActiveControl.Name
        Case "登録"
            Call CEkey.SetFs(T_1, True)
            MsgBox "該当データをセットします。登録処理は行いません。"
            Exit Sub

        Case T_1
            Call B_SET_Click
    End Select
'
End Sub

'------------------------------------------------
' 年月日_GotFocus
'------------------------------------------------
Private Sub 年月日_GotFocus()
    Call CEkey.AllSelect
End Sub

Private Sub T_1_LostFocus()
    T_1 = Right$(P8.FFormat(T_1, "#,##0"), 15)
End Sub

Private Sub T_2_LostFocus()
    T_2 = Right$(P8.FFormat(T_2, "#,##0"), 15)
'
    Select Case wFname2
    Case "内入入力"
        wd当月残高 = wd前月残高 - P8.FCDbl(T_2)
        L_4.Caption = Format(wd当月残高, "#,##0")
    End Select
'
End Sub

Private Sub T_3_LostFocus()
    T_3 = Right$(P8.FFormat(T_3, "#,##0"), 15)
End Sub

'------------------------------------------------
' 返済残高_セット
'------------------------------------------------
Private Sub 返済残高_セット()
'
    Dim wi01 As Integer
    Dim w融資残高 As Double
'
    '融資残高
    w融資残高 = wdYkin
    wi01 = 1
    
    w初回 = Null
    w最終 = Null
    
    wstr = ""
    wstr = wstr + "Select"
    wstr = wstr + " 実際年月日,元金額,融資残高,返済回数,取消フラグ"
    wstr = wstr + " From " & wsTbl
    wstr = wstr + " Where 借入番号 = '" & wsBango & "'"
    wstr = wstr + " And 取消フラグ=0"
    wstr = wstr + " Order by 実際年月日"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    Do Until wRs.eof
        wRs("融資残高") = w融資残高 - wRs("元金額")
        wRs("返済回数") = wi01
            
        If wi01 = 1 Then
            w初回 = wRs("実際年月日")
        End If
        
        w融資残高 = wRs("融資残高")
        wi01 = wi01 + 1
        
        wRs.Update
    
        wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
'
    '合計融資残高
    L_合計融資残高.Caption = ""
    
    wstr = ""
    wstr = wstr + "Select 実際年月日,融資残高"
    wstr = wstr + " From " & wsTbl
    wstr = wstr + " Where 借入番号 = '" & wsBango & "'"
    wstr = wstr + " And 取消フラグ=0"
    wstr = wstr + " Order by 実際年月日 desc"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
        L_合計融資残高.Caption = Format(P8.FCDbl(wRs("融資残高")), "#,##0")
        w最終 = wRs("実際年月日")
    End If
    
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
返済残高_セット_ERR:
    pERR_MES = pPROGRAM_ID + "/ 返済残高_セット() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

Private Sub L_利率_Click()

End Sub

'------------------------------------------------
' 削除_Click
'------------------------------------------------
Private Sub 削除_Click()
'
    Dim wi01 As Integer
    Dim wd01 As Double, wd02 As Double
    Dim wDate1 As Date, wDate2 As Date
'
    On Error GoTo 削除_Click_ERR
'
    ' =========================================
    '           権限チェック
    ' =========================================
    Select Case GUserKen
        Case "0"
            '入力権限
        Case "1"
            '照会権限
            MsgBox "権限がありません", vbExclamation
            Exit Sub
        Case "5"
            '管理者権限
        Case Else
            MsgBox "権限がありません", vbExclamation
            Exit Sub
    End Select

    GRet = MsgBox("削除しますよろしいですか？", vbYesNo + vbExclamation)
    If GRet = vbNo Then
        Exit Sub
    End If
'
    Select Case wFname2
    Case "内入入力"
        Call 削除_Uchiire
            
        Exit Sub
    End Select
'
    ' =========================================
    '            更新処理
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月日", 年月日.Text)
    If GVar1 = 0 Or GVar1 = Null Then
        Exit Sub
    End If
'
    ' =========================================
    '            明細TR
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月日", 年月日.Text)
    If GVar1 = 0 Or GVar1 = Null Then
        Exit Sub
    End If
    
    wstr = ""
    wstr = wstr + "Delete * From " & wsTbl
    wstr = wstr + " Where 借入番号 = '" & wsBango & "'"
    wstr = wstr + " And Format(実際年月日,'yyyy/mm/dd') = '" & Format(GVar1, "yyyy/mm/dd") & "'"
    GDb.Execute wstr
'
    Call 返済残高_セット
'
    '----------< DataGrid Close >----------------------------------------------
    Set DataGrid1.DataSource = Nothing
    Adodc1.Recordset.Close
'
    ' =========================================
    '               画面セット
    ' =========================================
    年月日.Text = ""
    
    Call 画面セット(False)
'    Call 登録後初期セット
    Call CEkey.SetFs(年月日, True)
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "削除しました。", vbInformation
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
削除_Click_ERR:
    pERR_MES = pPROGRAM_ID + "/ 削除_Click() でエラー" + vbCrLf + vbCrLf + _
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
' 登録_Click
'------------------------------------------------
Private Sub 登録_Click()
'
    Dim wi01 As Integer
    Dim wd01 As Double, wd02 As Double
    Dim wDate1 As Date, wDate2 As Date
'
    On Error GoTo 登録_Click_ERR
'

    ' =========================================
    '           権限チェック
    ' =========================================
    Select Case GUserKen
        Case "0"
            '入力権限
        Case "1"
            '照会権限
            MsgBox "権限がありません", vbExclamation
            Exit Sub
        Case "5"
            '管理者権限
        Case Else
            MsgBox "権限がありません", vbExclamation
            Exit Sub
    End Select
    
    Select Case wFname2
    Case "内入入力"
        Call 登録_Uchiire
            
        Exit Sub
    End Select
'
    '----------------------------------------
    '               登録チェック
    '----------------------------------------
    If Not IsNumeric(T_1) And T_1 <> "" Then
        MsgBox "入力を確認してください": Call CEkey.SetFs(T_1, True)
        Exit Sub
    End If
    
    If Not IsNumeric(T_2) And T_2 <> "" Then
        MsgBox "入力を確認してください": Call CEkey.SetFs(T_2, True)
        Exit Sub
    End If
'
    ' =========================================
    '            更新処理
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月日", 年月日.Text)
    If GVar1 = 0 Or GVar1 = Null Then
        Exit Sub
    End If
    
    '実行日
    wDate1 = C年月日.平成To西暦("年月日", L_実行日.Caption)
    wDate2 = C年月日.平成To西暦("年月日", 年月日.Text)
    If wDate1 > wDate2 Then
        MsgBox "初回返済年月が違います"
        Call CEkey.SetFs(年月日, True)
        Exit Sub
    End If
'
    ' =========================================
    '            明細TR
    ' =========================================
    wd01 = P8.FCDbl(T_1)
    wd02 = P8.FCDbl(T_2)
    L_4.Caption = Format(wd01 + wd02, "#,##0")
'
    GVar1 = C年月日.平成To西暦("年月日", 年月日.Text)
    If GVar1 = 0 Or GVar1 = Null Then
        Exit Sub
    End If
    
    wstr = ""
    wstr = wstr + "Select"
    wstr = wstr + " 借入番号,実際年月日,返済予定年月,元金額,利息額,返済金額,融資残高,"
    wstr = wstr + " 返済回数,取消フラグ"
    wstr = wstr + " From " & wsTbl
    wstr = wstr + " Where 借入番号 = '" & wsBango & "'"
    wstr = wstr + " And Format(実際年月日,'yyyy/mm/dd') = '" & Format(GVar1, "yyyy/mm/dd") & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.eof Then
            wRs.AddNew
            
            wRs("借入番号") = wsBango
            wRs("実際年月日") = CDate(GVar1)
        End If
     
            wRs("返済予定年月") = CDate(Format(GVar1, "yyyy/mm") & "/01")
            wRs("元金額") = wd01
            wRs("利息額") = wd02
            wRs("返済金額") = wd01 + wd02
            wRs("取消フラグ") = 0 'P8.FCDbl(取消)
            
            If wRs("取消フラグ") = 1 Then
                wRs("返済回数") = 0
            End If
            
        wRs.Update
    wRs.Close
    Set wRs = Nothing
'
    Call 返済残高_セット
'
    ' =========================================
    '            借入金、貸付金
    ' =========================================
    wstr = ""
    wstr = wstr + "Update " + wsTbl2
    
    If P8.FCDbl(L_合計融資残高.Caption) = 0 Then
        wstr = wstr + " Set 手入力区分=1"
    Else
        wstr = wstr + " Set 手入力区分=2"
    End If
    
    If Not IsNull(P8.FCDate(w初回)) Then
        wstr = wstr + " ,初回返済年月 =#" & Format(P8.FCDate(w初回), "yyyy/mm/dd") & "#,"
        wstr = wstr + " 初回返済実行日 =#" & Format(P8.FCDate(w初回), "yyyy/mm/dd") & "#"
    End If
    If Not IsNull(P8.FCDate(w最終)) Then
        wstr = wstr + " ,最終返済年月 =#" & Format(P8.FCDate(w最終), "yyyy/mm/dd") & "#,"
        wstr = wstr + " 最終返済実行日 =#" & Format(P8.FCDate(w最終), "yyyy/mm/dd") & "#"
    End If
    wstr = wstr + " Where 借入番号='" & wsBango & "'"
    GDb.Execute wstr
'
    '----------< DataGrid Close >----------------------------------------------
    Set DataGrid1.DataSource = Nothing
    Adodc1.Recordset.Close
'
    ' =========================================
    '               画面セット
    ' =========================================
    Call 画面セット(False)
'    Call 登録後初期セット
    Call CEkey.SetFs(年月日, True)
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "登録しました。", vbInformation
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
登録_Click_ERR:
    pERR_MES = pPROGRAM_ID + "/ 登録_Click() でエラー" + vbCrLf + vbCrLf + _
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
' 内入残高_セット
'------------------------------------------------
Private Sub 内入残高_セット(wdate As Date)
'
    On Error GoTo 内入残高_セット_ERR
'
    wd前月残高 = 0
    wd当月残高 = 0
    
    wd前月残高 = MBD010_借入前月残高(w借入データ, wdate)
    wd当月残高 = wd前月残高 - P8.FCDbl(T_2)
    
    L_1.Caption = Format(wd前月残高, "#,##0")
    L_4.Caption = Format(wd当月残高, "#,##0")
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
内入残高_セット_ERR:
    pERR_MES = pPROGRAM_ID + "/ 内入残高_セット() でエラー" + vbCrLf + vbCrLf + _
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
' 登録_Uchiire
'------------------------------------------------
Private Sub 登録_Uchiire()
'
    Dim wRs1 As ADODB.Recordset
    Dim wstr1 As String
    
    Dim j As Integer, k As Integer
    Dim wiCnt As Integer
    Dim FLG_DEL As Boolean
    Dim w年月日 As Variant, wv01 As Variant
    Dim ws01 As String
'
    On Error GoTo 登録_Uchiire_ERR
'
    FLG_DEL = False
    
    '----------------------------------------
    '               登録チェック
    '----------------------------------------
    w年月日 = C年月日.平成To西暦("年月", 年月日, True)
    
    If FLG_MAX = True Then
    
        For k = 1 To 500 '内入回数
            j = j + 1
            
            If IsNull(w借入内入.内入(k).内入x回目年月日) Or P8.FCStr(w借入内入.内入(k).内入x回目年月日) = "" Then
                Exit For
            End If
            
            If IsDate(w借入内入.内入(k).内入x回目年月日) Then
                If w借入内入.内入(k).内入x回目年月日 = w年月日 Then
                    Exit For
                End If
            End If
        
        Next k
                
        If k > 500 Then
            GRet = MsgBox("内入回数500回を越えると登録できません。", vbOKOnly)
            Call CEkey.SetFs(年月日, True)
            Exit Sub
        End If
    End If
'
    If P8.FCStr(年月日) = "" Then
        Exit Sub
    End If
    
    If C年月日.平成To西暦("年月", 年月日, True) = 0 Then
        MsgBox "年月日が違います"
        Call CEkey.SetFs(年月日, True)
        Exit Sub
    End If
'
    ' =========================================
    '             年月日整合性check
    ' =========================================
    If Not IsNull(w年月日) Then
        If CDate(w年月日) < CDate(wv実行日) Then
            MsgBox "年月日が誤りです"
            Call CEkey.SetFs(年月日, True)
            Exit Sub
        End If
        'If CDate(w年月日) < CDate(wv初回返済実行日) Or CDate(w年月日) > CDate(wv最終返済実行日) Then
        If CDate(w年月日) >= CDate(wv最終返済実行日) Then
            MsgBox "年月日が誤りです"
            Call CEkey.SetFs(年月日, True)
            Exit Sub
        End If
    End If
'
    If Not IsNull(wv解約実行日) Then
        If CDate(w年月日) = CDate(wv解約実行日) Then
            MsgBox "年月日が誤りです"
            Call CEkey.SetFs(年月日, True)
            Exit Sub
        End If
    End If
'
    ' =========================================
    '             内入金額check
    ' =========================================
    If wd前月残高 < P8.FCDbl(T_2) Or wd当月残高 < 0 Then
        MsgBox "内入金額を確認してください"
        Call CEkey.SetFs(T_2, True)
        Exit Sub
    End If
'
    ' =========================================
    '            更新処理
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月", 年月日)
    If GVar1 = 0 Then
        GVar1 = Null
    End If
'
    ' =========================================
    '             解約処理
    ' =========================================
    If wd前月残高 - P8.FCDbl(T_2) = 0 _
    And wd当月残高 = 0 Then
        GRet = MsgBox("期日前一括償還の処理になります" _
                & vbLf & "借入金登録の解約日に入力年月日をセットして、" _
                & vbLf & "解約処理をおこないます" & vbLf & "よろしいですか？", vbYesNo + vbQuestion)
        If GRet = vbYes Then
            Call 解約処理
                
                Exit Sub
        Else
            Call CEkey.SetFs(T_2, True)
            Exit Sub
        End If
    End If
'
    'If P8.FCDbl(取消) = 1 Then
    '    FLG_DEL = True
    'End If
'
    '----------< テーブル Write >----------------------------------------------
    If FLG_DEL <> True Then
        wstr = ""
        wstr = wstr + "Select * From  DCHA010_Gridワーク"
        wstr = wstr + " Where Format(年月日1,'yyyymmdd') = '" & Format(GVar1, "yyyymmdd") & "'"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.eof Then
            wRs.AddNew
            
            wRs("年月日1") = P8.FCDate(GVar1)
        
            wRs("数値1") = P8.FCDbl(T_2)
            wRs("数値2") = P8.FCDbl(T_3)
            
            '内入区分
            wRs("数値4") = wi据置X回目
            
            wRs.Update
        Else
        
            wRs("数値1") = P8.FCDbl(T_2)
            wRs("数値2") = P8.FCDbl(T_3)
            
            '内入区分
            wRs("数値4") = wi据置X回目
            
            wRs.Update
            
        End If
        
        wRs.Close
        Set wRs = Nothing
    Else
        wstr = ""
        wstr = wstr + "Delete * From  DCHA010_Gridワーク"
        wstr = wstr + " Where Format(年月日1,'yyyymmdd') = '" & Format(GVar1, "yyyymmdd") & "'"
        GDb.Execute wstr
    End If
'
    '----------< テーブル削除 >------------------------------------------------
    wstr = "Delete * from DBDA010_借入金内入1"
    wstr = wstr & " Where 借入番号='" & wsBango & "'"
    GDb.Execute wstr
    
    wstr = "Delete * from DBDA010_借入金内入2"
    wstr = wstr & " Where 借入番号='" & wsBango & "'"
    GDb.Execute wstr
    
    wstr = "Delete * from DBDA010_借入金内入3"
    wstr = wstr & " Where 借入番号='" & wsBango & "'"
    GDb.Execute wstr
    
    wstr = "Delete * from DBDA010_借入金内入4"
    wstr = wstr & " Where 借入番号='" & wsBango & "'"
    GDb.Execute wstr
    
    wstr = "Delete * from DBDA010_借入金内入5"
    wstr = wstr & " Where 借入番号='" & wsBango & "'"
    GDb.Execute wstr
    
    wstr = "Delete * from DBDA010_借入金内入6"
    wstr = wstr & " Where 借入番号='" & wsBango & "'"
    GDb.Execute wstr
    
    wstr = "Delete * from DBDA010_借入金内入7"
    wstr = wstr & " Where 借入番号='" & wsBango & "'"
    GDb.Execute wstr
    
    '----------< テーブル Write >----------------------------------------------
    wiCnt = 1
    
    wstr = ""
    wstr = wstr & "Select * From  DCHA010_Gridワーク"
    wstr = wstr & " where 数値4=1 or 数値4=3 or 数値4=4"
    wstr = wstr & " order by 年月日1"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    Do Until wRs.eof
    
            wstr1 = ""
            If wiCnt <= 80 Then
            'DBDA010_借入金内入1 1～80
                wstr1 = "Select * from DBDA010_借入金内入1"
            ElseIf wiCnt >= 81 And wiCnt <= 160 Then
            'DBDA010_借入金内入2 81～160
                wstr1 = "Select * from DBDA010_借入金内入2"
            ElseIf wiCnt >= 161 And wiCnt <= 240 Then
            'DBDA010_借入金内入3 161～240
                wstr1 = "Select * from DBDA010_借入金内入3"
            ElseIf wiCnt >= 241 And wiCnt <= 320 Then
            'DBDA010_借入金内入4 241～320
                wstr1 = "Select * from DBDA010_借入金内入4"
            ElseIf wiCnt >= 321 And wiCnt <= 400 Then
            'DBDA010_借入金内入5 321～400
                wstr1 = "Select * from DBDA010_借入金内入5"
            ElseIf wiCnt >= 401 And wiCnt <= 480 Then
            'DBDA010_借入金内入6 401～480
                wstr1 = "Select * from DBDA010_借入金内入6"
            ElseIf wiCnt >= 481 And wiCnt <= 500 Then
            'DBDA010_借入金内入7 481～500
                wstr1 = "Select * from DBDA010_借入金内入7"
            End If
            
            wstr1 = wstr1 & " Where 借入番号='" & wsBango & "'"
            Call AdoRecordsetOpen(GDb, wRs1, wstr1)
            If wRs1.eof Then
                wRs1.AddNew
        
                wRs1("借入番号") = wsBango
                
            End If
                
                ws01 = "内入" & CStr(wiCnt) & "回目年月日"
                wRs1(ws01) = P8.FCDate(wRs("年月日1"))
        
                ws01 = "内入金額" & CStr(wiCnt) & "回目"
                wRs1(ws01) = P8.FCDbl(wRs("数値1"))
        
                ws01 = "手数料" & CStr(wiCnt) & "回目"
                wRs1(ws01) = P8.FCDbl(wRs("数値2"))
        
                wRs1.Update
            
                wiCnt = wiCnt + 1
            
            wRs1.Close
            Set wRs1 = Nothing
        
        wRs.MoveNext
    Loop

    wRs.Close
    Set wRs = Nothing
'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    If 新規変更.Caption = "新規登録" Then
        wslog = "追加"
    ElseIf 新規変更.Caption = "変更" And 削除.Value = 0 Then
        wslog = "更新"
    End If
    GLogStr = "借入番号=" & wsBango & ","
    GLogStr = GLogStr & "年月日=" & Format(GVar1, "yyyy/mm/dd") & ","
    GLogStr = GLogStr & "内入金額=" & P8.FCDbl(T_2) & ","
    GLogStr = GLogStr & "手数料=" & P8.FCDbl(T_3)
    Call MXA030_LOG_WRITE("内入入力", wslog, GLogStr)
'
    ' =========================================
    '                 初期設定
    ' =========================================
    'ワークテーブル作成とワークデータセット
    Call 内入ワークテーブル作成
    
    Call 登録後初期セット
        
    Call CEkey.SetFs(年月日, False)
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "更新処理は終了しました", vbInformation
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
登録_Uchiire_ERR:
    pERR_MES = pPROGRAM_ID + "/ 登録_Uchiire() でエラー" + vbCrLf + vbCrLf + _
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
' 削除_Uchiire
'------------------------------------------------
Private Sub 削除_Uchiire()
'
    Dim wRs1 As ADODB.Recordset
    Dim wstr1 As String
    
    Dim j As Integer, k As Integer
    Dim wiCnt As Integer
    Dim FLG_DEL As Boolean
    Dim w年月日 As Variant, wv01 As Variant
    Dim ws01 As String
'
    On Error GoTo 削除_Uchiire_ERR
'
    FLG_DEL = False
    
    '----------------------------------------
    '               登録チェック
    '----------------------------------------
    If P8.FCStr(年月日) = "" Then
        Exit Sub
    End If
    
    If C年月日.平成To西暦("年月", 年月日, True) = 0 Then
        MsgBox "年月日が違います"
        Call CEkey.SetFs(年月日, True)
        Exit Sub
    End If
'
    ' =========================================
    '            更新処理
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月", 年月日)
    If GVar1 = 0 Then
        GVar1 = Null
    End If
'
    '----------< テーブル Write >----------------------------------------------
    wstr = ""
    wstr = wstr + "Delete * From  DCHA010_Gridワーク"
    wstr = wstr + " Where Format(年月日1,'yyyymmdd') = '" & Format(GVar1, "yyyymmdd") & "'"
    GDb.Execute wstr
'
    '----------< テーブル削除 >------------------------------------------------
    wstr = "Delete * from DBDA010_借入金内入1"
    wstr = wstr & " Where 借入番号='" & wsBango & "'"
    GDb.Execute wstr
    
    wstr = "Delete * from DBDA010_借入金内入2"
    wstr = wstr & " Where 借入番号='" & wsBango & "'"
    GDb.Execute wstr
    
    wstr = "Delete * from DBDA010_借入金内入3"
    wstr = wstr & " Where 借入番号='" & wsBango & "'"
    GDb.Execute wstr
    
    wstr = "Delete * from DBDA010_借入金内入4"
    wstr = wstr & " Where 借入番号='" & wsBango & "'"
    GDb.Execute wstr
    
    wstr = "Delete * from DBDA010_借入金内入5"
    wstr = wstr & " Where 借入番号='" & wsBango & "'"
    GDb.Execute wstr
    
    wstr = "Delete * from DBDA010_借入金内入6"
    wstr = wstr & " Where 借入番号='" & wsBango & "'"
    GDb.Execute wstr
    
    wstr = "Delete * from DBDA010_借入金内入7"
    wstr = wstr & " Where 借入番号='" & wsBango & "'"
    GDb.Execute wstr
    
    '----------< テーブル Write >----------------------------------------------
    wiCnt = 1
    
    wstr = ""
    wstr = wstr & "Select * From  DCHA010_Gridワーク"
    wstr = wstr & " where 数値4=1 or 数値4=3 or 数値4=4"
    wstr = wstr & " order by 年月日1"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    Do Until wRs.eof
    
            wstr1 = ""
            If wiCnt <= 80 Then
            'DBDA010_借入金内入1 1～80
                wstr1 = "Select * from DBDA010_借入金内入1"
            ElseIf wiCnt >= 81 And wiCnt <= 160 Then
            'DBDA010_借入金内入2 81～160
                wstr1 = "Select * from DBDA010_借入金内入2"
            ElseIf wiCnt >= 161 And wiCnt <= 240 Then
            'DBDA010_借入金内入3 161～240
                wstr1 = "Select * from DBDA010_借入金内入3"
            ElseIf wiCnt >= 241 And wiCnt <= 320 Then
            'DBDA010_借入金内入4 241～320
                wstr1 = "Select * from DBDA010_借入金内入4"
            ElseIf wiCnt >= 321 And wiCnt <= 400 Then
            'DBDA010_借入金内入5 321～400
                wstr1 = "Select * from DBDA010_借入金内入5"
            ElseIf wiCnt >= 401 And wiCnt <= 480 Then
            'DBDA010_借入金内入6 401～480
                wstr1 = "Select * from DBDA010_借入金内入6"
            ElseIf wiCnt >= 481 And wiCnt <= 500 Then
            'DBDA010_借入金内入7 481～500
                wstr1 = "Select * from DBDA010_借入金内入7"
            End If
            
            wstr1 = wstr1 & " Where 借入番号='" & wsBango & "'"
            Call AdoRecordsetOpen(GDb, wRs1, wstr1)
            If wRs1.eof Then
                wRs1.AddNew
        
                wRs1("借入番号") = wsBango
                
            End If
                
                ws01 = "内入" & CStr(wiCnt) & "回目年月日"
                wRs1(ws01) = P8.FCDate(wRs("年月日1"))
        
                ws01 = "内入金額" & CStr(wiCnt) & "回目"
                wRs1(ws01) = P8.FCDbl(wRs("数値1"))
        
                ws01 = "手数料" & CStr(wiCnt) & "回目"
                wRs1(ws01) = P8.FCDbl(wRs("数値2"))
        
                wRs1.Update
            
                wiCnt = wiCnt + 1
            
            wRs1.Close
            Set wRs1 = Nothing
        
        wRs.MoveNext
    Loop

    wRs.Close
    Set wRs = Nothing
'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    GLogStr = "借入番号=" & wsBango & ","
    GLogStr = "年月日=" & Format(GVar1, "yyyy/mm/dd")
    Call MXA030_LOG_WRITE("内入入力", "削除", GLogStr)
'
    ' =========================================
    '                 初期設定
    ' =========================================
    'ワークテーブル作成とワークデータセット
    Call 内入ワークテーブル作成
    
    Call 登録後初期セット
        
    Call CEkey.SetFs(年月日, False)
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "削除処理は終了しました", vbInformation
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
削除_Uchiire_ERR:
    pERR_MES = pPROGRAM_ID + "/ 削除_Uchiire() でエラー" + vbCrLf + vbCrLf + _
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
' 解約処理
'------------------------------------------------
Private Sub 解約処理()
'
    Dim wDate1 As Date
'
    On Error GoTo 解約処理_ERR
'
    ' =========================================
    '            更新処理
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月", 年月日)
    If GVar1 = 0 Then
        GVar1 = Null
    End If
'
    ' =========================================
    '            借入金、貸付金
    ' =========================================
    wstr = ""
    wstr = wstr + "Select 支払日,営業日区分,解約年月,解約実行日 From " + wsTbl2
    wstr = wstr + " Where 借入番号='" & wsBango & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
        wDate1 = MXA030_実行支払年月(CDate(GVar1), P8.FCDbl(wRs("支払日")), P8.FCDbl(wRs("営業日区分")), "")
        
        wRs("解約年月") = Format(wDate1, "yyyy/mm/dd")
        wRs("解約実行日") = Format(P8.FCDate(GVar1), "yyyy/mm/dd")
        
        wRs.Update
        
    End If
    
    wRs.Close
    Set wRs = Nothing
'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    wslog = "更新"
    GLogStr = "借入番号=" & wsBango & ","
    GLogStr = GLogStr & "解約日=" & Format(GVar1, "yyyy/mm/dd")
    Call MXA030_LOG_WRITE("内入入力", wslog, GLogStr)
'
    ' =========================================
    '                 初期設定
    ' =========================================
    'ワークテーブル作成とワークデータセット
    Call 内入ワークテーブル作成
    
    Call 登録後初期セット
        
    Call CEkey.SetFs(年月日, False)
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "更新処理は終了しました", vbInformation
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
解約処理_ERR:
    pERR_MES = pPROGRAM_ID + "/ 解約処理() でエラー" + vbCrLf + vbCrLf + _
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
' 閉じる_Click
'------------------------------------------------
Private Sub 閉じる_Click()
'
    '----------< DataGrid Close >----------------------------------------------
    If Not DataGrid1.DataSource Is Nothing Then
        Set DataGrid1.DataSource = Nothing
    End If
    Adodc1.Recordset.Close
'
    GStr = wFname
    GStr_1 = wsBango
    
    Unload Me
    
    frm_I借入金登録.Enabled = True
    Call frm_I借入金登録.画面セット呼出
'
End Sub


