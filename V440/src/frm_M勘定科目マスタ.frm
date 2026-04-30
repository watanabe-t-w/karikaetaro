VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_M勘定科目マスタ 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "勘定科目マスタ"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11295
   Icon            =   "frm_M勘定科目マスタ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10185
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox 削除データを表示 
      Caption         =   "削除データを表示"
      Height          =   255
      Left            =   6240
      TabIndex        =   30
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton CSV出力 
      Caption         =   "CSV出力"
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
      Left            =   360
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   9480
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "登録"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   240
      TabIndex        =   16
      Top             =   5640
      Width           =   10935
      Begin VB.ComboBox C_仕訳補助 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   2
         Top             =   1440
         Width           =   3255
      End
      Begin VB.CheckBox 貸方補助科目使用 
         Height          =   255
         Left            =   7920
         TabIndex        =   11
         Top             =   3120
         Width           =   495
      End
      Begin VB.CheckBox 借方補助科目使用 
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   3120
         Width           =   495
      End
      Begin VB.CheckBox 社債フラグ 
         Height          =   255
         Left            =   2040
         TabIndex        =   1
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox 仕訳補助備考 
         Enabled         =   0   'False
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
         IMEMode         =   1  'ｵﾝ
         Left            =   5160
         MaxLength       =   100
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox 仕訳補助 
         Enabled         =   0   'False
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
         IMEMode         =   3  'ｵﾌ固定
         Left            =   4440
         MaxLength       =   20
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox 貸方勘定科目名 
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
         IMEMode         =   1  'ｵﾝ
         Left            =   7440
         MaxLength       =   100
         TabIndex        =   10
         Top             =   2760
         Width           =   3255
      End
      Begin VB.TextBox 貸方勘定科目 
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
         IMEMode         =   1  'ｵﾝ
         Left            =   7440
         MaxLength       =   20
         TabIndex        =   9
         Top             =   2400
         Width           =   3255
      End
      Begin VB.TextBox 借方勘定科目名 
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
         IMEMode         =   1  'ｵﾝ
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   7
         Top             =   2760
         Width           =   3255
      End
      Begin VB.TextBox 借方勘定科目 
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
         IMEMode         =   1  'ｵﾝ
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   6
         Top             =   2400
         Width           =   3255
      End
      Begin VB.TextBox 仕訳名 
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
         IMEMode         =   1  'ｵﾝ
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1800
         Width           =   3255
      End
      Begin VB.CheckBox 削除 
         Caption         =   "削除"
         Height          =   255
         Left            =   5640
         TabIndex        =   17
         Top             =   1800
         Width           =   1695
      End
      Begin 借換たろう.ZU070_Label 新規変更 
         Height          =   375
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
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
      Begin 借換たろう.ZU020_ComboBox 仕訳区分 
         Height          =   345
         Left            =   1920
         TabIndex        =   0
         Top             =   720
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   609
         ForeColor       =   -2147483640
         ForeColor       =   -2147483640
         IMEMode         =   3
         TextWidth       =   615
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
      Begin VB.Label Label5 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 仕訳補助"
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
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label11 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "貸方補助科目使用"
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
         Left            =   5640
         TabIndex        =   27
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label10 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "借方補助科目使用"
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
         Left            =   120
         TabIndex        =   26
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label9 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 社債フラグ"
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
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 貸方勘定科目名"
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
         Left            =   5640
         TabIndex        =   24
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 貸方勘定科目"
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
         Left            =   5640
         TabIndex        =   23
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 借方勘定科目名"
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
         Left            =   120
         TabIndex        =   22
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 借方勘定科目"
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
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 仕訳区分名"
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
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   " 仕訳区分"
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
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   1815
      End
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
      Left            =   9360
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9480
      Width           =   1815
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
      Left            =   7440
      TabIndex        =   12
      Top             =   9480
      Width           =   1815
   End
   Begin 借換たろう.ZU050_Button ZU050_Button1 
      Height          =   615
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "勘定科目マスタ"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4725
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8334
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
      Left            =   2400
      Top             =   9480
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
End
Attribute VB_Name = "frm_M勘定科目マスタ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "勘定科目マスタ"

Dim wRs As ADODB.Recordset
Dim wstr As String
'
'------------------------------------------------
' Form_Initialize
'------------------------------------------------
'Private Sub Form_Initialize()
''
'    ' =========================================
'    '             MAA100_SERIAL
'    ' =========================================
'    GRet = MAA100_SERIAL()
'    If GRet <> True Then
'        GRet = MsgBox("シリアル情報が正しくありません。" + Chr(13) + vbCrLf + GProduct + "を終了します", vbOKOnly + vbCritical)
'        GDb.Close
'        Set GDb = Nothing
'
'        End
'    End If
''
'End Sub

'------------------------------------------------
' Form_Load
'------------------------------------------------
Private Sub Form_Load()
'
    ' =========================================
    '                 初期設定
    ' =========================================
'    Me.Caption = GFcap
    
    Me.Left = G_LEFT
    Me.Top = G_TOP
'
    With 仕訳区分
        .P8_SqlString = ""
        .P8_KeyLeng = 1
        
        Call .AddItem(1, "借入金(社債)の実行")
        Call .AddItem(2, "借入金(社債)の返済")
        Call .AddItem(3, "利息(社債)の支払")
        Call .AddItem(4, "利息(社債)の計上")
        Call .AddItem(5, "社債手数料の支払")
        Call .AddItem(6, "社債保証料の支払")
        Call .AddItem(7, "借入金(社債)の借入金長短振替")
    End With
    仕訳区分.CreateCombo
'
    With C_仕訳補助
        .Clear
        
        .AddItem "長短区分：長期借入金"
        .ItemData(C_仕訳補助.NewIndex) = XMXA020_区分("長短区分", "長期借入金")
        .AddItem "長短区分：短期借入金"
        .ItemData(C_仕訳補助.NewIndex) = XMXA020_区分("長短区分", "短期借入金")
        .AddItem "利息区分：利息先払"
        .ItemData(C_仕訳補助.NewIndex) = XMXA020_区分("利息区分", "利息先払")
        .AddItem "利息区分：利息後払"
        .ItemData(C_仕訳補助.NewIndex) = XMXA020_区分("利息区分", "利息後払")
        .AddItem "区分なし"
        .ItemData(C_仕訳補助.NewIndex) = "9"
    End With
'
    ' =========================================
    '                 初期設定
    ' =========================================
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
'
'
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
    GWhere = " Where (1=1) " + GWhere
    
    wstr = ""
    wstr = wstr & "Select "
    wstr = wstr & "仕訳区分 & 仕訳補助 & 社債フラグ As 仕訳番号,"
    wstr = wstr & "仕訳区分 As Grd仕訳区分,"
    wstr = wstr & "仕訳名 As Grd仕訳名,"
    wstr = wstr & "IIF(社債フラグ = 0,'','○') As Grd社債,"
    wstr = wstr & "社債フラグ,"
    wstr = wstr & "仕訳補助 As Grd仕訳補助,"
    
    wstr = wstr & "IIF(仕訳補助備考='長短区分',"
    wstr = wstr & "IIF(仕訳補助='0','長短区分：短期借入金',IIF(仕訳補助='1','長短区分：長期借入金','区分なし')),"
    wstr = wstr & "IIF(仕訳補助備考='利息区分',"
    wstr = wstr & "IIF(仕訳補助='1','利息区分：利息先払',IIF(仕訳補助='2','利息区分：利息後払','区分なし')),"
    wstr = wstr & "'区分なし')) As Grd仕訳補助備考,"
    
    wstr = wstr & "借方勘定科目 As Grd借方勘定科目,"
    wstr = wstr & "借方勘定科目名 As Grd借方勘定科目名,"
    wstr = wstr & "IIF(借方補助科目使用 <> 0,'○','×') As Grd借方補助科目使用,"
    'wstr = wstr & "IIF(借方個別補助科目使用 <> 0,'○','×') As Grd借方個別補助使用," '日本ガス仕様
    wstr = wstr & "貸方勘定科目 As Grd貸方勘定科目,"
    wstr = wstr & "貸方勘定科目名 As Grd貸方勘定科目名,"
    wstr = wstr & "IIF(貸方補助科目使用 <> 0,'○','×') As Grd貸方補助科目使用,"
    'wstr = wstr & "IIF(貸方個別補助科目使用 <> 0,'○','×') As Grd貸方個別補助使用," '日本ガス仕様
    wstr = wstr & "IIF(取消フラグ = 0,'','×') As Grd取消"
    wstr = wstr & " From DABA010_勘定科目マスタ"
    wstr = wstr & GWhere
    If Me.削除データを表示.Value = 0 Then
        wstr = wstr & " AND 取消フラグ = 0"
    End If
    wstr = wstr + " Order By 仕訳区分,社債フラグ,仕訳補助"
  
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        Call XZMA010_DataGrid_Set("仕訳名", "", 2200, "L")
        Call XZMA010_DataGrid_Set("社債", "", 550, "C")
        Call XZMA010_DataGrid_Set("仕訳補助備考", "補助名", 1900, "L")
        Call XZMA010_DataGrid_Set("借方勘定科目", "", 1050, "L")
        Call XZMA010_DataGrid_Set("借方勘定科目名", "", 1050, "L")
        Call XZMA010_DataGrid_Set("借方補助科目使用", "借補", 500, "L")
        'Call XZMA010_DataGrid_Set("借方個別補助使用", "借個", 500, "L") '日本ガス仕様
        Call XZMA010_DataGrid_Set("貸方勘定科目", "", 1050, "L")
        Call XZMA010_DataGrid_Set("貸方勘定科目名", "", 1050, "L")
        Call XZMA010_DataGrid_Set("貸方補助科目使用", "貸補", 500, "L")
        'Call XZMA010_DataGrid_Set("貸方個別補助使用", "貸個", 500, "L") '日本ガス仕様
        Call XZMA010_DataGrid_Set("取消", "", 550, "C")
    Call XZMA010_DataGrid_Action(DataGrid1)
  
'    メッセージ = ""
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
    Call CEkey.SetFs(仕訳区分, True)
End Sub

'------------------------------------------------
' DataGrid1_LostFocus
'------------------------------------------------
Private Sub DataGrid1_LostFocus()
'
    On Error Resume Next
        Dim wCheckValue As Variant
        wCheckValue = Adodc1.Recordset.Fields.Item("Grd仕訳区分")
        If Err.Number = 3021 Then GoTo Exit_Sub
    On Error GoTo Err_Hundle
        仕訳区分.Text = P8.FCStr(Adodc1.Recordset.Fields.Item("Grd仕訳区分"))
        仕訳補助 = P8.FCStr(Adodc1.Recordset.Fields.Item("Grd仕訳補助"))
        社債フラグ = P8.FCDbl(Adodc1.Recordset.Fields.Item("社債フラグ"))
    On Error GoTo 0
    
    Call 画面セット(True)
   
    If DataGrid1.Splits.Count <> 1 Then
        DataGrid1.Splits.Remove 1
    End If

    Call CEkey.SetFs(仕訳名, True)

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
    Call 画面セット(False)
    Call CEkey.AllSelect
End Sub

'------------------------------------------------
' 削除データを表示_Click
'------------------------------------------------
Private Sub 削除データを表示_Click()
    
    Call AdodcRefresh
    
End Sub

'------------------------------------------------
' 画面セット
'------------------------------------------------
Private Function 画面セット(pGridClick As Boolean) As Boolean
'
    On Error GoTo 画面セット_ERR
'
    画面セット = False
    
    ' =========================================
    '                画面クリア
    ' =========================================
    仕訳名 = ""
    '仕訳補助備考 = ""
    借方勘定科目.Text = ""
    借方勘定科目名.Text = ""
    貸方勘定科目.Text = ""
    貸方勘定科目名.Text = ""
    借方補助科目使用 = 0
    '借方個別補助使用 = 0 '日本ガス仕様
    貸方補助科目使用 = 0
    '貸方個別補助使用 = 0 '日本ガス仕様
    削除 = 0
    
    ' =========================================
    '            勘定科目マスタ セット
    ' =========================================
    wstr = ""
    wstr = wstr & "Select *"
    wstr = wstr & " From DABA010_勘定科目マスタ"
    wstr = wstr & " Where 仕訳区分 = '" & 仕訳区分.Text & "'"
    wstr = wstr & " And 仕訳補助 = '" & 仕訳補助 & "'"
    wstr = wstr & " And 社債フラグ = " & 社債フラグ
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.EOF Then
            If 仕訳区分.Text <> "" Then
                GRet = MsgBox("新規レコードを追加します。よろしいですか？", vbYesNo)
                If GRet = vbNo Then
                    新規変更.Caption = ""
                    wRs.Close
                    Set wRs = Nothing

                    Exit Function
                End If
                
                新規変更.Caption = "新規登録"
                Call CEkey.SetFs(仕訳名, True)
    
            End If
        Else
            画面セット = True
            
            Call CEkey.SetFs(仕訳名, True)
            新規変更.Caption = "変更"
            
            仕訳名 = P8.FCStr(wRs("仕訳名"))
            仕訳補助備考 = P8.FCStr(wRs("仕訳補助備考"))
            借方勘定科目 = P8.FCStr(wRs("借方勘定科目"))
            借方勘定科目名 = P8.FCStr(wRs("借方勘定科目名"))
            借方補助科目使用 = P8.FCDbl(wRs("借方補助科目使用"))
            '借方個別補助使用 = P8.FCDbl(wRs("借方個別補助科目使用")) '日本ガス仕様
            貸方勘定科目 = P8.FCStr(wRs("貸方勘定科目"))
            貸方勘定科目名 = P8.FCStr(wRs("貸方勘定科目名"))
            貸方補助科目使用 = P8.FCDbl(wRs("貸方補助科目使用"))
            '貸方個別補助使用 = P8.FCDbl(wRs("貸方個別補助科目使用")) '日本ガス仕様
            削除 = P8.FCDbl(wRs("取消フラグ"))

            If 仕訳補助備考 = "" Then
                C_仕訳補助.ListIndex = 4
            ElseIf 仕訳補助備考 = "長短区分" Then
                If 仕訳補助 = XMXA020_区分("長短区分", "長期借入金") Then
                    C_仕訳補助.ListIndex = 0
                ElseIf 仕訳補助 = XMXA020_区分("長短区分", "短期借入金") Then
                    C_仕訳補助.ListIndex = 1
                Else
                    C_仕訳補助.ListIndex = 4
                End If
            ElseIf 仕訳補助備考 = "利息区分" Then
                If 仕訳補助 = XMXA020_区分("利息区分", "利息先払") Then
                    C_仕訳補助.ListIndex = 2
                ElseIf 仕訳補助 = XMXA020_区分("利息区分", "利息後払") Then
                    C_仕訳補助.ListIndex = 3
                Else
                    C_仕訳補助.ListIndex = 4
                End If
            End If
        
        End If
    wRs.Close
    Set wRs = Nothing
    
    '------------------------------------------
    '          ** グリッドコントロール **
    '------------------------------------------
    If Not pGridClick Then
        DoEvents
        Call AdodcRefresh
    End If

    DoEvents
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "仕訳番号 = '" & 仕訳区分.Text & 仕訳補助 & 社債フラグ & "'")
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
' 検索_Click
'------------------------------------------------
Private Sub 検索_Click()
    Call 登録後初期セット
End Sub

Private Sub C_仕訳補助_Click()
'
    Dim j As Integer
'
    For j = 0 To C_仕訳補助.ListCount
        If C_仕訳補助.List(j) = C_仕訳補助 Then
            If C_仕訳補助.List(j) = "区分なし" Then
                仕訳補助 = 9
                仕訳補助備考 = ""
            ElseIf C_仕訳補助.List(j) = "長短区分：長期借入金" Then
                仕訳補助 = 1
                仕訳補助備考 = "長短区分"
            ElseIf C_仕訳補助.List(j) = "長短区分：短期借入金" Then
                仕訳補助 = 0
                仕訳補助備考 = "長短区分"
            ElseIf C_仕訳補助.List(j) = "利息区分：利息先払" Then
                仕訳補助 = 1
                仕訳補助備考 = "利息区分"
            ElseIf C_仕訳補助.List(j) = "利息区分：利息後払" Then
                仕訳補助 = 2
                仕訳補助備考 = "利息区分"
            Else
                仕訳補助 = 9
                仕訳補助備考 = ""
            End If
        End If
    Next j
'
End Sub

'------------------------------------------------
' 仕訳区分_GotFocus
'------------------------------------------------
Private Sub 仕訳区分_GotFocus()
    Call CEkey.AllSelect
End Sub

'------------------------------------------------
' 仕訳補助_LostFocus
'------------------------------------------------
Private Sub C_仕訳補助_LostFocus()
'
    On Error GoTo C_仕訳補助_LostFocus_ERR
'
    Select Case Screen.ActiveControl.Name
        Case "閉じる", "DataGrid1", "仕訳区分", "社債フラグ", "仕訳補助", "C_仕訳補助"
            Exit Sub
'        Case Else
'            Exit Sub
    End Select
   
    Call 画面セット(False)
    Call CEkey.AllSelect

    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
C_仕訳補助_LostFocus_ERR:
    pERR_MES = pPROGRAM_ID + "/ 仕訳区分_LostFocus() でエラー" + vbCrLf + vbCrLf + _
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
    Dim w仕訳区分 As String
'
    w仕訳区分 = 仕訳区分.Text
    
    仕訳区分.Text = ""
    Call 画面セット(False)
    新規変更.Caption = ""
    
    '----------------------------------------
    '               更新行を表示
    '----------------------------------------
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "仕訳番号 = '" & 仕訳区分.Text & 仕訳補助 & 社債フラグ & "'")
    Call CEkey.SetFs(仕訳区分, True)
'
End Sub

'------------------------------------------------
' LostFocus
'------------------------------------------------
Private Sub 仕訳名_LostFocus()
    Call P8.FCControlLeft(仕訳名, 20)
End Sub

'------------------------------------------------
' 登録_Click
'------------------------------------------------
Private Sub 登録_Click()
'
    Dim wslog As String
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
'
    If P8.FCStr(仕訳区分.Text) = "" Then
        MsgBox "仕訳区分が未入力です。", vbExclamation
        Call CEkey.SetFs(仕訳区分, True)
        Exit Sub
    End If

    'If P8.FCStr(仕訳補助) = "" Then
    '    MsgBox "仕訳補助が未入力です。", vbExclamation
    '    Call CEkey.SetFs(仕訳補助, True)
    '    Exit Sub
    'End If

    If 仕訳名 = "" Then
        MsgBox "仕訳名が未入力です。", vbExclamation
        Call CEkey.SetFs(仕訳名, True)
        Exit Sub
    End If
'
    If P8.FCStr(借方勘定科目) = "" Then
        MsgBox "借方勘定科目が未入力です。", vbExclamation
        Call CEkey.SetFs(借方勘定科目, True)
        Exit Sub
    End If

    If 借方勘定科目名 = "" Then
        MsgBox "借方勘定科目名が未入力です。", vbExclamation
        Call CEkey.SetFs(借方勘定科目名, True)
        Exit Sub
    End If
    
    If P8.FCStr(貸方勘定科目) = "" Then
        MsgBox "貸方勘定科目が未入力です。", vbExclamation
        Call CEkey.SetFs(貸方勘定科目, True)
        Exit Sub
    End If

    If 貸方勘定科目名 = "" Then
        MsgBox "貸方勘定科目名が未入力です。", vbExclamation
        Call CEkey.SetFs(貸方勘定科目名, True)
        Exit Sub
    End If
'
    'If Not IsNumeric(利息区分.Text) Or 利息区分.Text = "" Then
    '    MsgBox "利息区分を選択してください。", vbExclamation
    '    Call CEkey.SetFs(利息区分, True)
    '    Exit Sub
    'End If
'
    'If Not IsNumeric(長短区分.Text) Or 長短区分.Text = "" Then
    '    MsgBox "長短区分を選択してください。", vbExclamation
    '    Call CEkey.SetFs(長短区分, True)
    '    Exit Sub
    'End If
'
    ' =========================================
    '            勘定科目マスタ 更新処理
    ' =========================================
    wstr = ""
    wstr = wstr & "Select *"
    wstr = wstr & " From DABA010_勘定科目マスタ"
    wstr = wstr & " Where 仕訳区分 = '" & 仕訳区分.Text & "'"
    wstr = wstr & " And 仕訳補助 = '" & 仕訳補助 & "'"
    wstr = wstr & " And 社債フラグ = " & 社債フラグ
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.EOF Then
            wRs.AddNew
            
            wRs("仕訳区分") = 仕訳区分.Text
            wRs("仕訳補助") = 仕訳補助
            wRs("社債フラグ") = 社債フラグ
            
            wslog = "追加"
        End If
     
        wRs("仕訳名") = P8.FCStr(仕訳名.Text)
        wRs("仕訳補助備考") = P8.FCStr(仕訳補助備考.Text)
        wRs("借方勘定科目") = P8.FCStr(借方勘定科目.Text)
        wRs("借方勘定科目名") = P8.FCStr(借方勘定科目名.Text)
        wRs("借方補助科目使用") = P8.FCDbl(借方補助科目使用.Value)
        wRs("借方個別補助科目使用") = 0
        'wRs("借方個別補助科目使用") = P8.FCDbl(借方個別補助使用.Value) '日本ガス仕様
        wRs("貸方勘定科目") = P8.FCStr(貸方勘定科目.Text)
        wRs("貸方勘定科目名") = P8.FCStr(貸方勘定科目名.Text)
        wRs("貸方補助科目使用") = P8.FCDbl(貸方補助科目使用.Value)
        wRs("貸方個別補助科目使用") = 0
        'wRs("貸方個別補助科目使用") = P8.FCDbl(貸方個別補助使用.Value) '日本ガス仕様
        
        wRs("取消フラグ") = P8.FCDbl(削除.Value)
 
        wRs.Update
    wRs.Close
    Set wRs = Nothing
'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    If 新規変更.Caption = "新規" Then
        wslog = "追加"
    ElseIf 新規変更.Caption = "変更" And 削除.Value = 0 Then
        wslog = "更新"
    ElseIf 新規変更.Caption = "変更" And 削除.Value = 1 Then
        wslog = "削除"
    End If
    GLogStr = "仕訳区分=" & P8.FCStr(仕訳区分.Text) & ","
    GLogStr = GLogStr & "仕訳名=" & P8.FCStr(仕訳名.Text) & ","
    GLogStr = GLogStr & "仕訳補助=" & P8.FCStr(仕訳補助.Text) & ","
    GLogStr = GLogStr & "仕訳補助備考=" & P8.FCStr(仕訳補助備考.Text) & ","
    GLogStr = GLogStr & "社債フラグ=" & P8.FCStr(社債フラグ.Value) & ","
    GLogStr = GLogStr & "借方勘定科目=" & P8.FCStr(借方勘定科目.Text) & ","
    GLogStr = GLogStr & "借方勘定科目名=" & P8.FCStr(借方勘定科目名.Text) & ","
    GLogStr = GLogStr & "借方補助科目使用=" & P8.FCStr(借方補助科目使用.Value) & ","
    'GLogStr = GLogStr & "借方個別補助科目使用=" & P8.FCStr(借方個別補助使用.Value) & "," '日本ガス仕様
    GLogStr = GLogStr & "貸方勘定科目=" & P8.FCStr(貸方勘定科目.Text) & ","
    GLogStr = GLogStr & "貸方勘定科目名=" & P8.FCStr(貸方勘定科目名.Text) & ","
    GLogStr = GLogStr & "貸方補助科目使用=" & P8.FCStr(貸方補助科目使用.Value) & ","
    'GLogStr = GLogStr & "貸方個別補助科目使用=" & P8.FCStr(貸方個別補助使用.Value) & "," '日本ガス仕様
    GLogStr = GLogStr & "削除=" & P8.FCStr(削除.Value)
    Call MXA030_LOG_WRITE(pPROGRAM_ID, wslog, GLogStr)
'
    Adodc1.Refresh
'
    ' =========================================
    '               画面セット
    ' =========================================
    Call 画面セット(False)
    Call CEkey.SetFs(仕訳名, True)
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "登録しました", vbInformation
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
' CSV出力_Click
'------------------------------------------------
Private Sub CSV出力_Click()
'
    Call MX040_勘定科目(GKeyName & "_" & "勘定科目.csv")
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
    Unload Me
End Sub
