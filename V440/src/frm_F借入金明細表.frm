VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_F借入金明細表 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "借入金明細表　照会"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12870
   Icon            =   "frm_F借入金明細表.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   12870
   Begin VB.TextBox 金利変更年月1 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   330
      IMEMode         =   2  'ｵﾌ
      Left            =   6600
      TabIndex        =   31
      Text            =   "HH年MM月"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox 解約シミュレーション 
      Caption         =   "解約シミュレーション"
      Height          =   255
      Left            =   2160
      TabIndex        =   30
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton 登録データ照会 
      Caption         =   "登録データ照会"
      Height          =   495
      Left            =   7080
      TabIndex        =   4
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CheckBox 金利SM 
      Caption         =   "金利シミュレーション"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   3000
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6165
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3360
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   10874
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   13
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
         Size            =   8.25
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
      Top             =   0
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
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame 金利変更登録フレーム 
      Caption         =   "金利変更登録"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   9120
      TabIndex        =   19
      Top             =   840
      Width           =   3615
      Begin VB.CommandButton B_SET 
         Caption         =   "SET"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton 登録 
         Caption         =   "登録"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   1920
         Width           =   855
      End
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
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox 金利変更利率 
         Alignment       =   1  '右揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   2  'ｵﾌ
         Left            =   1680
         TabIndex        =   1
         Text            =   "00.00000"
         Top             =   1440
         Width           =   1455
      End
      Begin 借換たろう.ZU070_Label 新規変更 
         Height          =   375
         Left            =   240
         TabIndex        =   28
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
      Begin VB.Label 金利変更年月 
         Alignment       =   1  '右揃え
         BorderStyle     =   1  '実線
         Caption         =   "HH年MM月DD日"
         Height          =   330
         Left            =   1680
         TabIndex        =   32
         Top             =   720
         Width           =   1425
      End
      Begin VB.Label Label57 
         Alignment       =   2  '中央揃え
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3120
         TabIndex        =   24
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label34 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "利率"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label L_金利変更年月日 
         Alignment       =   1  '右揃え
         BorderStyle     =   1  '実線
         Caption         =   "HH年MM月DD日"
         Height          =   330
         Left            =   1680
         TabIndex        =   22
         Top             =   1080
         Width           =   1425
      End
      Begin VB.Label Label31 
         Alignment       =   1  '右揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "変更年月日"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label30 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00D6DBBD&
         BorderStyle     =   1  '実線
         Caption         =   "変更年月"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   1455
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
      Left            =   10560
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   240
      Width           =   2175
   End
   Begin 借換たろう.ZU050_Button ZU050_Button1 
      Height          =   615
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "借入金明細表 照会"
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
   Begin VB.Label Label45 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  '実線
      Caption         =   " 銀行番号"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   26
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label H_銀行番号 
      Alignment       =   2  '中央揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      TabIndex        =   25
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label H_日割計算区分 
      Alignment       =   2  '中央揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      Caption         =   "自動計算"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      TabIndex        =   18
      Top             =   2550
      Width           =   1575
   End
   Begin VB.Label H_登録方法 
      Alignment       =   2  '中央揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      Caption         =   "標準登録"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      TabIndex        =   17
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label H_銀行名 
      Alignment       =   2  '中央揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      TabIndex        =   16
      Top             =   1950
      Width           =   4455
   End
   Begin VB.Label H_借入内容 
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      TabIndex        =   15
      Top             =   1380
      Width           =   3495
   End
   Begin VB.Label H_借入番号 
      Alignment       =   2  '中央揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      TabIndex        =   14
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label H_借入金種別 
      Alignment       =   2  '中央揃え
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      TabIndex        =   13
      Top             =   1110
      Width           =   2295
   End
   Begin VB.Label Label22 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "日割計算区分"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   12
      Top             =   2550
      Width           =   1575
   End
   Begin VB.Label Label21 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "登録方法"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label L_借入内容 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   " 借入内容"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   9
      Top             =   1380
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  '実線
      Caption         =   " 銀行名"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   10
      Top             =   1950
      Width           =   1575
   End
   Begin VB.Label L_借入番号 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   " 借入番号"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label28 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   " 借入金種別"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   7
      Top             =   1110
      Width           =   1575
   End
End
Attribute VB_Name = "frm_F借入金明細表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "借入明細表　照会"

Dim wRs As ADODB.Recordset
Dim wstr As String

Dim FLG_New As Boolean
Dim FLG_TR As Boolean

Dim w円単位 As String, wsTbl As String, wsTbl2 As String

'金利変更------------------------------------------------------------------------------------------------
Dim wi利息支払 As Integer, wi支払日 As Integer, wi営業日区分 As Integer, wi金利種別 As Integer
Dim wi利息計算日数区分 As Integer, wi返済単位 As Integer
Dim ws利息区分 As String
Dim wFname As String ', wsTbl As String
Dim wsBango As String

Dim wv初回返済実行日 As Variant, wv最終返済実行日 As Variant
Dim wv実行日 As Variant, wv初回返済年月 As Variant, wv最終返済年月 As Variant, wv初回金利年月 As Variant
Dim FLG_MAX As Boolean
Dim wslog As String

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
'
'------------------------------------------------
' Form_Load
'------------------------------------------------
Private Sub Form_Load()
'
    ' =========================================
    '                 入力モード
    ' =========================================
    Select Case G借入明細表照会.入力モード
        Case "0"
        '入力なし
            Me.金利変更登録フレーム.Visible = False

        Case "1"
        '金利変更入力
            Me.金利変更登録フレーム.Visible = True
            Me.新規変更.Caption = ""
            FLG_MAX = False
            
            wsBango = G借入明細表照会.借入番号
            
            金利変更年月 = ""
            L_金利変更年月日.Caption = ""
            金利変更利率 = 0
            '取消 = 0
            
            wv初回返済実行日 = Null
            wv最終返済実行日 = Null
            
            'ワークテーブル作成とワークデータセット
            Call 金利ワークテーブル作成
            
'            Call 金利変更画面セット
        

        Case "2"

        Case Else
            Me.金利変更登録フレーム.Visible = False
    End Select
    
    

    ' =========================================
    '                 初期設定
    ' =========================================
    If G借入明細表照会.金融リストラ番号 = "" Or G借入明細表照会.金融解約日 = "" Then
        解約シミュレーション.Enabled = False
    Else
        解約シミュレーション.Enabled = True
    End If
    
    If G借入明細表照会.金利シミュレーションGP = "" Then
        金利SM.Enabled = False
    Else
        金利SM.Enabled = True
    End If
    
    Me.Left = G_LEFT
    Me.Top = G_TOP
    
    G金利SM = False
    
    Call 登録後初期セット

End Sub

'------------------------------------------------
' Form_Activate
'------------------------------------------------
Private Sub Form_Activate()
'
    DoEvents
'
    ' =========================================
    '             コンボボックス
    ' =========================================
'
    Call CEkey.AllSelect
'
End Sub

'------------------------------------------------
' Form_KeyDown
'------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF11 Then
'        Call 登録_Click
'    End If
'
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
    Dim wRs As ADODB.Recordset
    Dim wWhere As String
    
    Dim w借入データ As MAA910_借入金
    Dim wdHRiritu As Double
    Dim w金融リストラ As String
    
    Dim w借入番号 As String
    
    Dim wdKinri As Double
'
    On Error GoTo AdodcRefresh_ERR
'
    ' =========================================
    '             グッリドの初期値
    ' =========================================
'    Call MXA030_DataGridInit(DataGrid1)
    DataGrid1.AllowRowSizing = False
    DataGrid1.HeadFont.Size = 11
    DataGrid1.HeadFont.Bold = True
    DataGrid1.Font.Size = 10
    
    If 金利SM.Value = 1 Or 解約シミュレーション.Value = 1 Then
        DataGrid1.BackColor = C_LGreen
    Else
        DataGrid1.BackColor = C_Yellow
    End If
    DataGrid1.ForeColor = RGB(0, 0, 160)

    Set DataGrid1.DataSource = Adodc1
  
    ' =========================================
    '              ConnectionString
    ' =========================================
    Call AdodcSet(Adodc1, GDb)
  
    ' =========================================
    '              メインクエリ
    ' =========================================
'
    '----------------------------------------------------------------
    '                         ** 初期設定 **
    '----------------------------------------------------------------
    wsTbl = "DBDA010_借入金"
    wsTbl2 = "DBDA010_借入金明細TR"
'
    '通常 or 手入力 の書式設定は↓
'
    '----------------------------------------------------------------
    '                           ** パラメータセット **
    '----------------------------------------------------------------
    w借入番号 = G借入明細表照会.借入番号
    w金融リストラ = G借入明細表照会.金融リストラ番号
    
    If 解約シミュレーション.Value = 1 Then
    Else
        w金融リストラ = ""
    End If
    
    '----------------------------------------------------------------
    '          ** グループ ・ Where ・ OrderＢｙ **
    '----------------------------------------------------------------
    wWhere = ""
    
    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    FLG_TR = False
    
    '** 明細ファイル 削除 **
    wstr = ""
    wstr = wstr + "Delete * From DCDA020_借入金明細"
    GDb.Execute wstr
    
    '** 明細ファイル 作成 **
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From " & wsTbl & " As k"
    wstr = wstr + " Where K.借入番号 = '" & P8.FCStr(w借入番号) + "'"
    wstr = wstr + wWhere
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    '手入力の場合は借入金データセットしない
'    If P8.FCDbl(wRs("手入力区分")) = "0" Then
      Do Until wRs.eof
      
          w借入データ = MBD010_借入データセット(wRs)
          If P8.FCDbl(wRs("手入力区分")) = "0" Then
          '標準
              Call MBD010_借入金テーブル作成(w金融リストラ, w借入データ)
              Call MBD010_借入明細作成(w金融リストラ, w借入データ)        ' 07/02/21 V180
          Else
          '入力登録
            Call MBD010_借入金入力明細Read(w借入データ)
            If w借入データ.社債フラグ = 1 Then
                Call MDA020_借入金入力社債明細作成(w借入データ)
            End If
            Call MBD010_借入明細作成_入力登録(w借入データ)
              
              FLG_TR = True
          End If

          wRs.MoveNext
      Loop
    wRs.Close
    Set wRs = Nothing
'
    '----------------------------------------------------------------
    '                          ** 合計 計算 **
    '----------------------------------------------------------------
    
    '----------------------------------------------------------------
    '                       ** レコード　ソース **
    '----------------------------------------------------------------
    wstr = ""
    wstr = wstr + "Select "
'    wstr = wstr + " K.借入計画番号 As H00_借入計画番号,"
    wstr = wstr + " K.金融リストラ番号 As 金融リストラ番号,"
    wstr = wstr + " IIF(K.sm区分=0,'OFF','ON') As sm区分,"
    wstr = wstr + " K.借入番号 As 借入番号,"
    wstr = wstr + " K.借入内容 As 借入内容,"
    
    wstr = wstr + " Format(K.実行日,'" & Gfmt年月 & "') As 実行日,"
    wstr = wstr + " Format(K.初回返済年月,'" & Gfmt年月 & "') As 初回返済年月,"
    wstr = wstr + " Format(K.初回返済実行日,'" & Gfmt年月日 & "') As 初回返済年月日,"
    wstr = wstr + " Format(K.最終返済年月,'" & Gfmt年月 & "') As 最終返済年月,"
    wstr = wstr + " Format(K.最終返済実行日,'" & Gfmt年月日 & "') As 最終返済年月日,"
    wstr = wstr + " Format(K.金利初回年月,'" & Gfmt年月 & "') As 金利初回年月,"
    wstr = wstr + " Format(K.解約実行日,'" & Gfmt年月日 & "') As 解約年月日,"
    wstr = wstr + " Format(K.金融解約実行日,'" & Gfmt年月日 & "') As 金融解約日,"

    wstr = wstr + " K.融資金額 As 融資金額,"
    wstr = wstr + " KSM.金利グループ名 AS 金利グループ名,"
    
    wstr = wstr + " K.保証料率 As 保証料率,"
    
    wstr = wstr + " KS.借入金種別名 AS 借入種別,"
    
    '手入力の場合は借入金データセットしない
    'HederとDetailの所で表示制御
    If FLG_TR <> True Then
        
        wstr = wstr + " Format(K.支払回数,'#,##0') As 支払回数,"
        wstr = wstr + " Format(K.据置回数,'#,##0') As 据置回数,"
        wstr = wstr + " K.返済単位月数 As 返済単位月数,"
        
        '変動金利の場合
        If P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) = w借入データ.金利種別 Then
            If w借入データ.変動最終利率 > -1 Then
                wstr = wstr + " '*' As 利率フラグ,"
                wstr = wstr + "'" & w借入データ.変動最終利率 & "' As 利率,"
            Else
                wstr = wstr + " '' As 利率フラグ,"
                wstr = wstr + " K.利率 As 利率,"
            End If
        Else
            wstr = wstr + " K.利率 As 利率,"
            wstr = wstr + " '' As 利率フラグ,"
        End If
    Else
        
        wstr = wstr + " Format(" & w借入データ.支払回数 & ",'#,##0') As 支払回数,"
        If P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) = w借入データ.金利種別 Then
            If w借入データ.変動最終利率 > -1 Then
                wstr = wstr + " '' As 利率フラグ,"
                wstr = wstr + "'" & w借入データ.変動最終利率 & "' As 利率,"
            Else
                wstr = wstr + " '' As 利率フラグ,"
                wstr = wstr + " K.利率 As 利率,"
            End If
        Else
            wstr = wstr + " K.利率 As 利率,"
            wstr = wstr + " '' As 利率フラグ,"
        End If
    End If
    
    wstr = wstr + " G.銀行番号 As 銀行番号,"
    wstr = wstr + " G.銀行名 As 銀行名,"
    wstr = wstr + " K.支払日,"
    wstr = wstr + " S.支払区分名 As 支払区分,"
    wstr = wstr + " IIF(K.営業日区分 = " & P8.FCDbl(XMXA020_区分("営業日", "翌営業日")) & ",'翌営業日','前営業日') As 営業日,"
    wstr = wstr + " IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As 利息区分,"
    wstr = wstr + " IIF(K.利息計算日数区分 = " & P8.FCDbl(XMXA020_区分("利息計算日数", "営業日数")) & ",'営業日数','固定日数') As 利息日数,"
    wstr = wstr + " IIF(K.利息支払方法 = " & P8.FCDbl(XMXA020_区分("利息支払", "毎月")) & ",'毎月','一括') As 利息支払方法,"
    wstr = wstr + " IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "控除無し")) & ",'控除無し',"
    wstr = wstr + " IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "実行日控除")) & ",'実行日控除',"
    wstr = wstr + " IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "最終返済日控除")) & ",'最終返済日控除',"
    wstr = wstr + " IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) & ",'実行日及び最終返済日控除','中間利払最終日控除')))) As 利息控除区分,"
    wstr = wstr + " IIF(K.金利計算年間日数 = " & P8.FCDbl(XMXA020_区分("金利計算", "365日")) & ",'365日','360日') As 金利年間日数,"
    wstr = wstr + " IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As 金利種別,"
    wstr = wstr + " K.金利条件 As 金利条件,"
    wstr = wstr + " IIF(K.有担保フラグ=0,'無担保','有担保') As 担保区分,"
    wstr = wstr & " IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期借入金','長期借入金') AS H00_長短区分,"
    wstr = wstr + " IIF(K.設備フラグ=0,'運転資金','設備') As H00_設備区分,"
    
    wstr = wstr + " K.借入金種別区分 As 借入金種別区分,"
    wstr = wstr + " H.保証会社区分名 As 保証会社区分名,"
    wstr = wstr + " Y.融資区分名 As 融資区分名,"
    
    wstr = wstr + " KM.据置X回目 As I_据置X回目,"
    
    wstr = wstr + " Format(KM.返済回数,'#,##0') As Grd返済回数,"
'    wstr = wstr + " Format(KM.日割日数,'#,##0') As Grd日割日数,"
    wstr = wstr + " Format(KM.利息対象期間日数,'#,##0') As Grd調整日数,"
    wstr = wstr + " Format(KM.実際年月日,'" & Gfmt年月日 & "') As Grd返済年月日,"
    wstr = wstr + " Format(KM.返済予定年月,'" & Gfmt年月日 & "') As Grd返済予定年月,"
    wstr = wstr + " Format(KM.利息計算年月日,'" & Gfmt年月日 & "') As Grd利息計算年月日,"
    wstr = wstr + " Format(KM.返済金額,'###,###,###,##0') As Grd返済金額,"
    wstr = wstr + " Format(KM.元金額,'###,###,###,##0') As Grd元金額,"
    wstr = wstr + " Format(KM.利息額,'###,###,###,##0') As Grd利息額,"
    wstr = wstr + " Format(KM.仮計上利息額,'###,###,###,##0') As Grd調整利息額,"
    wstr = wstr + " Format(KM.融資残高,'#,###,###,###,##0') As Grd融資残高,"
    wstr = wstr + " Format(KM.日割日数,'#,##0') As Grd日割日数,"
    wstr = wstr + " KM.手数料 As Grd手数料,"
    wstr = wstr + " Format(KM.利率,'#0.00000') As Grd利率,"
    wstr = wstr + " IIF(KM.返済予定年月 = WK.年月日1,'*','') As Grd金利変更フラグ"
    
    If w借入データ.社債フラグ = 1 Then
        wstr = wstr + ","
        wstr = wstr + " Format(KM.初期手数料,'###,###,###,##0') As Grd初期手数料,"
        wstr = wstr + " Format(KM.元金手数料,'###,###,###,##0') As Grd元金手数料,"
        wstr = wstr + " Format(KM.利息手数料,'###,###,###,##0') As Grd利息手数料,"
        wstr = wstr + " Format(KM.保証料,'###,###,###,##0') As Grd保証料"
    End If
    
    wstr = wstr + " From (((((((( DCDA020_借入金明細  As KM"
    wstr = wstr + " Inner Join " & wsTbl & " As K"
    wstr = wstr + "  ON KM.借入番号 = K.借入番号)"
    wstr = wstr + " Inner Join DAAA040_銀行マスタ As G"
    wstr = wstr + "  ON K.銀行番号 = G.銀行番号)"
    wstr = wstr + " Inner Join DAAB020_支払区分マスタ As S"
    wstr = wstr + "  ON K.支払日 = S.支払日)"
    wstr = wstr + " Left Join DCHA010_Gridワーク As WK"
    wstr = wstr + "  ON KM.借入番号 = WK.テキスト1 And KM.返済予定年月 = WK.年月日1)"
    wstr = wstr + " Left Join DAAA115_金利シミュレーショングループ As KSM"
    wstr = wstr + "  ON K.金利グループ区分 = KSM.金利グループ区分)"
    wstr = wstr + " Left Join DAAA116_借入金種別 As KS"
    wstr = wstr + "  ON K.借入金種別区分 = KS.借入金種別区分)"
    wstr = wstr + " Left Join DAAA100_保証会社区分 As H"
    wstr = wstr + "  ON K.保証会社区分 = H.保証会社区分)"
    wstr = wstr + " Left Join DAAA110_融資区分 As Y"
    wstr = wstr + "  ON K.融資区分 = Y.融資区分)"
'    wstr = wstr + " lnner Join DCHA010_Gridワーク As WK"
'    wstr = wstr + "  ON KM.借入番号 = WK.テキスト1)"
    
    wstr = wstr + " Order BY KM.実際年月日,KM.据置X回目"

    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
    'ヘッダー情報のセット
        
        Me.H_銀行番号 = P8.FCStr(wRs("銀行番号"))
        Me.H_銀行名 = P8.FCStr(wRs("銀行名"))
        Me.H_借入番号 = P8.FCStr(wRs("借入番号"))
        Me.H_借入金種別 = P8.FCStr(wRs("借入種別"))
        Me.H_借入内容 = P8.FCStr(wRs("借入内容"))
        
        If FLG_TR = True Then
            Me.H_登録方法 = "入力登録"
        Else
            Me.H_登録方法 = "標準登録"
        End If
        Me.H_日割計算区分 = ""
        
    End If
    
    Adodc1.RecordSource = wstr
    Adodc1.Refresh

    Call XZMA010_DataGrid_Init
        Call XZMA010_DataGrid_Set("返済回数", "回", 500, "C")
        Call XZMA010_DataGrid_Set("返済予定年月", "", 0, "R")
        Call XZMA010_DataGrid_Set("返済年月日", "返済日", 1500, "R")
        Call XZMA010_DataGrid_Set("利息計算年月日", "利息計算日", 1500, "R")
        Call XZMA010_DataGrid_Set("元金額", "返済元金", 1600, "R")
        Call XZMA010_DataGrid_Set("利息額", "支払利息", 1600, "R")
        Call XZMA010_DataGrid_Set("返済金額", "支払合計", 1600, "R")
        Call XZMA010_DataGrid_Set("融資残高", "融資残高", 1700, "R")
        Call XZMA010_DataGrid_Set("日割日数", "日数", 700, "R")
        Call XZMA010_DataGrid_Set("利率", "利率", 1000, "R")
        Call XZMA010_DataGrid_Set("金利変更フラグ", " ", 300, "C")
        
        If w借入データ.社債フラグ = 1 Then
            Call XZMA010_DataGrid_Set("初期手数料", "", 1500, "R")
            Call XZMA010_DataGrid_Set("元金手数料", "", 1500, "R")
            Call XZMA010_DataGrid_Set("利息手数料", "", 1500, "R")
            Call XZMA010_DataGrid_Set("保証料", "", 1500, "R")
        End If
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
' DataGrid1_Click
'------------------------------------------------
Private Sub DataGrid1_Click()
'
'    Call CEkey.SetFs(基準金利コード, True)

    ' =========================================
    '                 入力モード
    ' =========================================
    Select Case G借入明細表照会.入力モード
        Case "0"
        '入力なし
            Exit Sub
        Case "1"
        '金利変更入力
            Call CEkey.SetFs(金利変更利率, True)

        Case "2"
        '明細入力
            Me.金利変更登録フレーム.Visible = False

        Case Else
            Exit Sub
    End Select
    
End Sub

'------------------------------------------------
' DataGrid1_LostFocus
'------------------------------------------------
Private Sub DataGrid1_LostFocus()
'

    If G借入明細表照会.入力モード = 0 Then
        Exit Sub
    End If
    
    On Error Resume Next
        Dim wCheckValue As Variant
        wCheckValue = Adodc1.Recordset.Fields.Item("Grd返済予定年月")
        If Err.Number = 3021 Then GoTo Exit_Sub
    On Error GoTo Err_Hundle
        金利変更年月 = C年月日.FormatDate("年月", P8.FCStr(Adodc1.Recordset.Fields.Item("Grd返済予定年月")))
        金利変更利率 = Format(P8.FCDbl(Adodc1.Recordset.Fields.Item("Grd利率")), "#0.00000")
    On Error GoTo 0
    
    L_金利変更年月日 = C年月日.FormatDate("年月日", P8.FCStr(Adodc1.Recordset.Fields.Item("Grd利息計算年月日")))
    
    Call 金利変更画面セット
   
'    If DataGrid1.Splits.Count <> 1 Then
'        DataGrid1.Splits.Remove 1
'    End If

'    Call CEkey.SetFs(基準金利名, True)

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
    Dim p借入番号 As String
'
    On Error GoTo 画面セット_ERR
'
    画面セット = False
    
    ' =========================================
    '                画面クリア
    ' =========================================
    
    ' =========================================
    '                パラメータ
    ' =========================================
    p借入番号 = P8.FCStr(G借入明細表照会.借入番号)
    
    ' =========================================
    '            ユーザーマスタ セット
    ' =========================================
    
'
'
    '------------------------------------------
    '          ** グリッドコントロール **
    '------------------------------------------
    If Not pGridClick Then
        DoEvents
        Call AdodcRefresh
    End If

    DoEvents
'    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd基準金利区分 = '" + 基準金利コード + "'")
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
' 基準金利コード_GotFocus
'------------------------------------------------
Private Sub 基準金利コード_GotFocus()
    Call CEkey.AllSelect
End Sub

'------------------------------------------------
' 登録後初期セット
'------------------------------------------------
Private Sub 登録後初期セット()
'
    Dim w基準金利コード As String
'
'    w基準金利コード = 基準金利コード
    
'    基準金利コード = ""
    Call 画面セット(False)
    
    '----------------------------------------
    '               更新行を表示
    '----------------------------------------
'    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd基準金利区分 = '" + w基準金利コード + "'")
'    Call CEkey.SetFs(Me.基準金利コード, True)
'
End Sub

Private Sub 基準金利コード_LostFocus()

    Call 画面セット(False)
    
End Sub


'------------------------------------------------
' 入力クリア_Click
'------------------------------------------------
Private Sub 入力クリア_Click()
    Call 登録後初期セット

End Sub

Private Sub 削除データを表示_Click()
    
    Call AdodcRefresh
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frm_I借入金登録.Enabled = True
End Sub



Private Sub 解約シミュレーション_Click()

    Select Case 金利SM.Value
        Case 0: G金利SM = False
        Case 1: G金利SM = True
    End Select
    
    Call AdodcRefresh

End Sub

Private Sub 金利SM_Click()
    
    Select Case 金利SM.Value
        Case 0: G金利SM = False
        Case 1: G金利SM = True
    End Select
    
    Call AdodcRefresh
    
End Sub

Private Sub 西暦表示_Click()

    Select Case 金利SM.Value
        Case 0: G金利SM = False
        Case 1: G金利SM = True
    End Select
    
    Call AdodcRefresh

End Sub

Private Sub 登録データ照会_Click()
    
    frm_F借入登録データ照会.Show
    
End Sub

'------------------------------------------------
' 閉じる_Click
'------------------------------------------------
Private Sub 閉じる_Click()
'
    Unload frm_F借入登録データ照会
    Unload Me

End Sub


'********************************************************************************************************
'------------------------------------------------
'金利変更登録
'------------------------------------------------
'********************************************************************************************************

'------------------------------------------------
' 金利ワークテーブル作成
'------------------------------------------------
Private Sub 金利ワークテーブル作成()
'
    Dim wRs1 As ADODB.Recordset
    Dim wstr1 As String
    
    Dim j As Integer
    Dim ws01 As String
'
    On Error GoTo 金利ワークテーブル作成_ERR
'
    wsTbl = "DBDA010_借入金"

    '----------< ワークテーブル削除 >------------------------------------------
    wstr = "Delete * from DCHA010_Gridワーク"
    GDb.Execute wstr
'
    If wsBango = "" Then
'    If G借入明細表照会.借入番号 = "" Then
        Exit Sub
    End If
'
    wi支払日 = 0
    wi営業日区分 = 0
    wi金利種別 = 0
    
    wv初回返済実行日 = Null
    wv最終返済実行日 = Null

    wi利息支払 = 0
    wi返済単位 = 1
    wi利息計算日数区分 = 0
    ws利息区分 = ""
    wv実行日 = Null
    wv初回返済年月 = Null
    wv最終返済年月 = Null
    wv初回金利年月 = Null
    
    '----------< テーブル Write >----------------------------------------------
    wstr = "Select * from DCHA010_Gridワーク"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        wstr1 = "Select * from " & wsTbl
        wstr1 = wstr1 & " Where 借入番号='" & wsBango & "'"
        Call AdoRecordsetOpen(GDb, wRs1, wstr1)
        If Not wRs1.eof Then
            
            wi支払日 = P8.FCDbl(wRs1("支払日"))
            wi営業日区分 = P8.FCDbl(wRs1("営業日区分"))
            wi金利種別 = P8.FCDbl(wRs1("金利種別"))
            wv初回返済実行日 = wRs1("初回返済実行日")
            wv最終返済実行日 = wRs1("最終返済実行日")
            
            wi利息支払 = P8.FCDbl(wRs1("利息支払方法"))
            wi返済単位 = P8.FCDbl(wRs1("返済単位月数"))
            wi利息計算日数区分 = P8.FCDbl(wRs1("利息計算日数区分"))
            ws利息区分 = P8.FCStr(wRs1("利息区分"))
            wv実行日 = wRs1("実行日")
            wv初回返済年月 = wRs1("初回返済年月")
            wv最終返済年月 = wRs1("最終返済年月")
            wv初回金利年月 = wRs1("金利初回年月")
            
            For j = 2 To 100
                
                ws01 = "金利変更" & CStr(j) & "回目年月"
                If Not IsNull(P8.FCDate(wRs1(ws01))) Then
                    
                    wRs.AddNew
                    
                    wRs("テキスト1") = wsBango
                    wRs("テキスト2") = j
                    
                    ws01 = "金利変更" & CStr(j) & "回目年月"
                    wRs("年月日1") = P8.FCDate(wRs1(ws01))
                    
                    ws01 = "金利" & CStr(j) & "回目"
                    wRs("数値1") = P8.FCDbl(wRs1(ws01))
                
                    wRs.Update
                    
                End If
                
            Next
            
            If Not IsNull(P8.FCDate(wRs1("金利変更１００回目年月"))) Then
                FLG_MAX = True
            End If
        
        End If
        wRs1.Close
        Set wRs1 = Nothing

    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
金利ワークテーブル作成_ERR:
    pERR_MES = pPROGRAM_ID + "/ 金利ワークテーブル作成() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

''------------------------------------------------
'' AdodcRefresh
''------------------------------------------------
'Private Sub AdodcRefresh()
''
'    On Error GoTo AdodcRefresh_ERR
''
'    ' =========================================
'    '             グッリドの初期値
'    ' =========================================
'    Call MXA030_DataGridInit(DataGrid1)
'    Set DataGrid1.DataSource = Adodc1
'
'    ' =========================================
'    '              ConnectionString
'    ' =========================================
'    Call AdodcSet(Adodc1, GDb)
'
'    ' =========================================
'    '              メインクエリ
'    ' =========================================
'    GWhere = ""
'    GWhere = " Where (1=1) " + GWhere
'
'    wstr = ""
'    wstr = wstr + "Select"
'    wstr = wstr + " テキスト2 As Grd回,"
'    wstr = wstr + " format(年月日1,Gfmt年月) As Grd年月,"
'    wstr = wstr + " format(数値1,'#,##0.00000') As Grd利率"
'    'wstr = wstr + " IIF(取消フラグ = 0,'','×') As Grd取消"
'    wstr = wstr + " From DCHA010_Gridワーク"
'    wstr = wstr + GWhere
'    wstr = wstr + " Order By 年月日1"
'
'    Adodc1.RecordSource = wstr
'    Adodc1.Refresh
'
'    Call XZMA010_DataGrid_Init
'        Call XZMA010_DataGrid_Set("回", "", 600, "L")
'        Call XZMA010_DataGrid_Set("年月", "金利変更年月", 2000, "R")
'        Call XZMA010_DataGrid_Set("利率", "金利変更利率", 2000, "R")
'        'Call XZMA010_DataGrid_Set("取消", "", 550, "C")
'    Call XZMA010_DataGrid_Action(DataGrid1)
'
'    メッセージ = ""
''
'    Exit Sub
''
''----------< ERROR ROUTINE >---------------------------------------------------
'AdodcRefresh_ERR:
'    pERR_MES = pPROGRAM_ID + "/ AdodcRefresh() でエラー" + vbCrLf + vbCrLf + _
'                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
'                "プロジェクト名：" + Err.Source + vbCrLf + _
'                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
'                GProduct + "を終了します"
'    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
'    pERR_RET = PUT_LOG(pERR_MES)
'
'    End
''
'End Sub

''------------------------------------------------
'' DataGrid1_Click
''------------------------------------------------
'Private Sub DataGrid1_Click()
''
'    メッセージ = ""
'    Call CEkey.SetFs(金利変更年月, True)
'End Sub
'
''------------------------------------------------
'' DataGrid1_LostFocus
''------------------------------------------------
'Private Sub DataGrid1_LostFocus()
''
'    On Error Resume Next
'        Dim wCheckValue As Variant
'        wCheckValue = Adodc1.Recordset.Fields.Item("Grd年月")
'        If Err.Number = 3021 Then GoTo Exit_Sub
'    On Error GoTo Err_Hundle
'        金利変更年月 = P8.FCStr(Adodc1.Recordset.Fields.Item("Grd年月"))
'    On Error GoTo 0
'
'    Call 画面セット
'
'    If DataGrid1.Splits.Count <> 1 Then
'        DataGrid1.Splits.Remove 1
'    End If
'
'    Call CEkey.SetFs(金利変更年月, True)
'
'Exit_Sub:
'    Exit Sub
'    '---------------------------------------------------
'Err_Hundle:
'    If Err.Number = 91 Then Resume Next
'    If Err.Number = 94 Then Resume Next
'    MsgBox CStr(Err.Number) + ":" + Err.Description
'    Resume Exit_Sub
'End Sub

'------------------------------------------------
' B_SET_Click
'------------------------------------------------
Private Sub B_SET_Click()
    Call 金利変更画面セット
    Call CEkey.AllSelect
End Sub

'------------------------------------------------
' 画面セット
'------------------------------------------------
Private Function 金利変更画面セット() As Boolean
'
    On Error GoTo 金利変更画面セット_ERR
'
    金利変更画面セット = False
'
'    金利変更利率 = 0
    '取消 = 0
    
    
    ' =========================================
    '                画面クリア
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月", 金利変更年月)
    If GVar1 = 0 Then
        GVar1 = Null
    End If
    
'    If Not IsNull(GVar1) And GVar1 <> "" Then
'        GRet = 金利変更年月CHECK(Format(GVar1, "yyyy/mm/dd"))
'        If GRet <> True Then
'            GRet = MsgBox("金利変更年月を確認してください", vbOKOnly + vbCritical)
'            Me.新規変更.Caption = ""
'            金利変更年月 = ""
'            L_金利変更年月日.Caption = ""
'            金利変更利率 = 0
'
'            Call CEkey.SetFs(金利変更年月, True)
'                Exit Function
'        End If
'    End If
    
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
            Me.新規変更.Caption = "新規"
            Me.削除.Enabled = False
            If FLG_MAX = True Then
                GRet = MsgBox("金利変更100回を越えると登録できません。", vbOKOnly)
                wRs.Close
                Set wRs = Nothing
                Me.新規変更.Caption = ""
                Exit Function
            End If
'            Call 金利変更年月日_セット
            Call CEkey.SetFs(金利変更利率, True)
        End If
    Else
        Me.新規変更.Caption = "変更"
        金利変更年月 = Format(wRs("年月日1"), Gfmt年月)
        金利変更利率 = P8.FFormat(P8.FCDbl(wRs("数値1")), "#,##0.00000")
        Me.削除.Enabled = True
'        Call 金利変更年月日_セット
    
    End If
    wRs.Close
    Set wRs = Nothing
    
    ' =========================================
    '            Grid セット
    ' =========================================
'    Call AdodcRefresh

    DoEvents
    Call XZMA010_DataGrid_Bookmark(DataGrid1, Adodc1, "Grd利息計算年月日= '" + C年月日.FormatDate("年月日", L_金利変更年月日) + "'")

'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
金利変更画面セット_ERR:
    pERR_MES = pPROGRAM_ID + "/ 金利変更画面セッ() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

Private Function 金利変更年月CHECK(pDate As Variant) As Boolean
'
    Dim wi01 As Integer
    Dim wd01 As Date
    Dim wvStr As Variant, wvEnd As Variant, wv01 As Variant
'
    On Error GoTo 金利変更年月CHECK_ERR
'
    金利変更年月CHECK = False
    
    If ws利息区分 = XMXA020_区分("利息区分", "利息先払") Then
        If CStr(wi利息支払) = XMXA020_区分("利息支払", "毎月") Then
            wvStr = wv初回金利年月
            wvEnd = DateAdd("m", -1, CDate(wv最終返済年月))
            
            If Format(wvStr, "yyyy/mm/01") <= Format(pDate, "yyyy/mm/01") _
            And Format(wvEnd, "yyyy/mm/01") >= Format(pDate, "yyyy/mm/01") Then
                金利変更年月CHECK = True
            End If
        
        ElseIf CStr(wi利息支払) = XMXA020_区分("利息支払", "一括") Then
            If Format(wv初回返済年月, "yyyy/mm/01") > Format(pDate, "yyyy/mm/01") Then
                wvStr = wv初回金利年月
                wvEnd = wv初回返済年月
                
                wv01 = wvStr
                Do While Format(wv01, "yyyy/mm/01") < Format(wvEnd, "yyyy/mm/01")
                    If Format(wv01, "yyyy/mm/01") = Format(pDate, "yyyy/mm/01") Then
                        金利変更年月CHECK = True
                        Exit Do
                    End If
                    wv01 = DateAdd("m", wi返済単位, CDate(wv01))
                Loop
            
            Else
                wvStr = wv初回返済年月
                wvEnd = DateAdd("m", -wi返済単位, CDate(wv最終返済年月))
                
                wv01 = wvStr
                Do While Format(wv01, "yyyy/mm/01") <= Format(wvEnd, "yyyy/mm/01")
                    If Format(wv01, "yyyy/mm/01") = Format(pDate, "yyyy/mm/01") Then
                        金利変更年月CHECK = True
                        Exit Do
                    End If
                    wv01 = DateAdd("m", wi返済単位, CDate(wv01))
                Loop
            End If
        End If

    ElseIf ws利息区分 = XMXA020_区分("利息区分", "利息後払") Then
        If CStr(wi利息支払) = XMXA020_区分("利息支払", "毎月") Then
            wvStr = DateAdd("m", 1, CDate(wv初回金利年月))
            wvEnd = wv最終返済年月
                
            If Format(wvStr, "yyyy/mm/01") <= Format(pDate, "yyyy/mm/01") _
            And Format(wvEnd, "yyyy/mm/01") >= Format(pDate, "yyyy/mm/01") Then
                金利変更年月CHECK = True
            End If
        
        ElseIf CStr(wi利息支払) = XMXA020_区分("利息支払", "一括") Then
        
            If Format(wv初回返済年月, "yyyy/mm/01") > Format(pDate, "yyyy/mm/01") Then
                wvStr = DateAdd("m", wi返済単位, CDate(wv初回金利年月))
                wvEnd = wv初回返済年月
                
                wv01 = wvStr
                Do While Format(wv01, "yyyy/mm/01") < Format(wvEnd, "yyyy/mm/01")
                    If Format(wv01, "yyyy/mm/01") = Format(pDate, "yyyy/mm/01") Then
                        金利変更年月CHECK = True
                        Exit Do
                    End If
                    wv01 = DateAdd("m", wi返済単位, CDate(wv01))
                Loop
            Else
                wvStr = DateAdd("m", wi返済単位, CDate(wv初回金利年月))
                wvEnd = wv最終返済年月
                
                wv01 = wvStr
                Do While Format(wv01, "yyyy/mm/01") <= Format(wvEnd, "yyyy/mm/01")
                    If Format(wv01, "yyyy/mm/01") = Format(pDate, "yyyy/mm/01") Then
                        金利変更年月CHECK = True
                        Exit Do
                    End If
                    wv01 = DateAdd("m", wi返済単位, CDate(wv01))
                Loop
            End If
            
        End If
    End If
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
金利変更年月CHECK_ERR:
    pERR_MES = pPROGRAM_ID + "/ 金利変更年月CHECK() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

Private Sub 金利変更年月_LostFocus()
    L_金利変更年月日.Caption = ""

    金利変更年月 = C年月日.FormatDate("年月", 金利変更年月)
    L_金利変更年月日 = ""
'
    If P8.FCStr(金利変更年月) <> "" Then
        Call 金利変更年月日_セット
    End If
End Sub

Private Sub 金利変更利率_LostFocus()
    金利変更利率 = P8.FFormat(金利変更利率, "#,##0.00000")
End Sub

'------------------------------------------------
' 金利変更年月日_セット
'------------------------------------------------
Private Sub 金利変更年月日_セット()
'
    Dim wv01 As Variant, wv02 As Variant
'
    On Error GoTo 金利変更年月日_セット_ERR
'
    GVar1 = C年月日.平成To西暦("年月", 金利変更年月)
    'GVar1 = MXA030_翌営業年月日計算(CDate(GVar1), wi支払日, wi営業日区分)
    wv01 = MBD010_利息計算年月日(CDate(GVar1), wi支払日, wi営業日区分, wi利息計算日数区分)
'
    If Format(GVar1, "yyyy/mm/01") = Format(wv初回返済年月, "yyyy/mm/01") Then
        If Format(wv01, "yyyy/mm/dd") = Format(wv初回返済実行日, "yyyy/mm/dd") Then
            wv01 = wv初回返済実行日
        End If
    End If
'
    If Format(GVar1, "yyyy/mm/01") = Format(wv最終返済年月, "yyyy/mm/01") Then
        If Format(wv01, "yyyy/mm/dd") = Format(wv最終返済実行日, "yyyy/mm/dd") Then
            wv01 = wv最終返済実行日
        End If
    End If
'
    L_金利変更年月日.Caption = Format(wv01, Gfmt年月日)
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
金利変更年月日_セット_ERR:
    pERR_MES = pPROGRAM_ID + "/ 金利変更年月日_セット() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

Private Sub 削除_Click()
'
    Dim wRs1 As ADODB.Recordset
    Dim wstr1 As String
    
    Dim j As Integer
    Dim FLG_DEL As Boolean
    Dim w金利変更年月日 As Variant, wv01 As Variant
    Dim ws01 As String
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

    FLG_DEL = False
    
    '----------------------------------------
    '               登録チェック
    '----------------------------------------
    GVar1 = C年月日.平成To西暦("年月", 金利変更年月)
    If GVar1 = 0 Or GVar1 = Null Then
        Exit Sub
    End If
    
    If C年月日.平成To西暦("年月", 金利変更年月, True) = 0 Then
        MsgBox "年月日が違います"
        Call CEkey.SetFs(金利変更年月, True)
        Exit Sub
    End If
'
    GRet = MsgBox("金利変更を取り消します。よろしいですか？", vbYesNo + vbExclamation)
    If GRet = vbNo Then
        Exit Sub
    End If
'
    ' =========================================
    '            更新処理
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月", 金利変更年月)
    If GVar1 = 0 Then
        GVar1 = Null
    End If
    
    FLG_DEL = True
    
    '----------< 取消データ削除 >------------------------------------------
    wstr = ""
    wstr = wstr + "Delete *"
    wstr = wstr + " From DCHA010_Gridワーク"
    wstr = wstr + " Where Format(年月日1,'yyyymmdd') = '" & Format(GVar1, "yyyymmdd") & "'"
    GDb.Execute wstr
'
    '----------< テーブル Write >----------------------------------------------
    wstr1 = "Select * from " & wsTbl
    wstr1 = wstr1 & " Where 借入番号='" & wsBango & "'"
    Call AdoRecordsetOpen(GDb, wRs1, wstr1)
    If Not wRs1.eof Then

        j = 2 '2回目から始まる
        
        wstr = "Select * from DCHA010_Gridワーク"
        wstr = wstr & " Where テキスト1='" & wsBango & "'"
        wstr = wstr & " Order by 年月日1"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        If Not wRs.eof Then
            Do Until wRs.eof
            
                ws01 = "金利変更" & CStr(j) & "回目年月"
                wRs1(ws01) = P8.FCDate(wRs("年月日1"))
    
                ws01 = "金利" & CStr(j) & "回目"
                wRs1(ws01) = P8.FCDbl(wRs("数値1"))
    
                j = j + 1
                
                wRs.MoveNext
            Loop
            
            If j <= 100 Then
                Do Until j > 100
                    ws01 = "金利変更" & CStr(j) & "回目年月"
                    wRs1(ws01) = Null
        
                    ws01 = "金利" & CStr(j) & "回目"
                    wRs1(ws01) = 0
        
                    j = j + 1
                Loop
            End If
        
            wRs1.Update
        
        Else
        
            If j <= 100 Then
                Do Until j > 100
                    ws01 = "金利変更" & CStr(j) & "回目年月"
                    wRs1(ws01) = Null
        
                    ws01 = "金利" & CStr(j) & "回目"
                    wRs1(ws01) = 0
        
                    j = j + 1
                Loop
            End If
        
            wRs1.Update
        
        End If
        
        wRs.Close
        Set wRs = Nothing
        
    End If
    wRs1.Close
    Set wRs1 = Nothing

    ' =========================================
    '               LOG_WRITE
    ' =========================================
    wslog = "削除"
    GLogStr = "金利変更登録:借入番号=" & P8.FCStr(Me.H_借入番号.Caption) & ","
    GLogStr = GLogStr & "年月日=" & P8.FCStr(Me.L_金利変更年月日)
    Call MXA030_LOG_WRITE(pPROGRAM_ID, wslog, GLogStr)

'
    ' =========================================
    '                 初期設定
    ' =========================================
    If FLG_DEL = True Then
        金利変更年月 = ""
        金利変更利率 = 0
    
        L_金利変更年月日.Caption = ""
    End If
'
    'ワークテーブル作成とワークデータセット
    Call 金利ワークテーブル作成
    
    Call 金利変更画面セット
        
    Call AdodcRefresh
    
    Call CEkey.SetFs(金利変更年月, False)
'
    ' =========================================
    '               メッセージ
    ' =========================================
'    メッセージ = "削除処理は終了しました"
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
    Dim wRs1 As ADODB.Recordset
    Dim wstr1 As String
    
    Dim j As Integer
    Dim FLG_DEL As Boolean
    Dim w金利変更年月日 As Variant, wv01 As Variant
    Dim ws01 As String
'
    On Error GoTo 保存_Click_ERR
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

    FLG_DEL = False
    
    '----------------------------------------
    '               登録チェック
    '----------------------------------------
    GVar1 = C年月日.平成To西暦("年月", 金利変更年月)
    If GVar1 = 0 Or GVar1 = Null Then
        MsgBox "金利変更年月が未入力です", vbExclamation
        Exit Sub
    End If
    
    GRet = 金利変更年月CHECK(Format(GVar1, "yyyy/mm/dd"))
    If GRet <> True Then
        MsgBox "指定された年月日での金利変更はできません", vbExclamation
        Call CEkey.SetFs(金利変更年月, True)
        Exit Sub
    End If
    
    If FLG_MAX = True Then
        GRet = MsgBox("金利変更100回を越えると登録できません。", vbOKOnly)
        Call CEkey.SetFs(金利変更年月, True)
        Exit Sub
    End If
'
    If C年月日.平成To西暦("年月", 金利変更年月, True) = 0 Then
        MsgBox "指定された年月日での金利変更はできません", vbExclamation
        Call CEkey.SetFs(金利変更年月, True)
        Exit Sub
    End If
'
    If (Not IsNumeric(金利変更利率) And 金利変更利率 <> "") _
    Or P8.FCDbl(金利変更利率) >= 100 Or P8.FCDbl(金利変更利率) < 0 Then
        MsgBox "入力を確認してください", vbExclamation
        Call CEkey.SetFs(金利変更利率, True)
        Exit Sub
    End If
'
    '固定金利で金利変更Ｘ回目年月等があればエラー
    If P8.FCDbl(XMXA020_区分("金利種別", "固定金利")) = wi金利種別 Then
        If 金利変更年月 <> "" Then
            MsgBox "固定金利では設定できません", vbExclamation
            Call CEkey.SetFs(金利変更年月, True)
            Exit Sub
        End If
        If P8.FCDbl(金利変更利率) <> 0 Then
            MsgBox "固定金利では設定できません", vbExclamation
            Call CEkey.SetFs(金利変更利率, True)
            Exit Sub
        End If
    End If
    
    If 金利変更年月 = "" Then
        If P8.FCDbl(金利変更利率) <> 0 Then
            MsgBox "金利変更利率が違います", vbExclamation
            Call CEkey.SetFs(金利変更年月, True)
            Exit Sub
        End If
    Else
        If IsNumeric(金利変更利率) Then
            If CInt(金利変更利率) > 100 Then
                MsgBox "金利変更利率が大きすぎます", vbExclamation
                Call CEkey.SetFs(金利変更利率, True)
                Exit Sub
            End If
            
            '2020/07/29 修正
            'If CInt(金利変更利率) = 0 Then
            If P8.FCDbl(金利変更利率) = 0 Then
                MsgBox "利率を確認してください", vbExclamation
                Call CEkey.SetFs(金利変更利率, True)
                Exit Sub
            End If
        End If
    End If
'
    ' =========================================
    '             金利変更年月整合性check
    ' =========================================
'    w金利変更年月日 = C年月日.平成To西暦("年月日", P8.FCStr(L_金利変更年月日.Caption), True)
'    If Not IsNull(w金利変更年月日) Then
'        If CDate(w金利変更年月日) < CDate(wv初回返済実行日) Or CDate(w金利変更年月日) > CDate(wv最終返済実行日) Then
'            MsgBox "金利変更年月が誤りです"
'            Call CEkey.SetFs(金利変更年月, True)
'            Exit Sub
'        End If
'    End If
'
    ' =========================================
    '            更新処理
    ' =========================================
    GVar1 = C年月日.平成To西暦("年月", 金利変更年月)
    If GVar1 = 0 Then
        GVar1 = Null
    End If
    
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DCHA010_Gridワーク"
    wstr = wstr + " Where Format(年月日1,'yyyymmdd') = '" & Format(GVar1, "yyyymmdd") & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.eof Then
        wRs.AddNew
    End If
        
        wRs("テキスト1") = wsBango
        
        wRs("年月日1") = C年月日.平成To西暦("年月", 金利変更年月, True)
        wRs("数値1") = P8.FCDbl(金利変更利率)
        
        wRs("取消フラグ") = 0 'P8.FCDbl(取消)
        
        'If P8.FCDbl(取消) = 1 Then
        '    FLG_DEL = True
        'End If
        
        wRs.Update
    
    wRs.Close
    Set wRs = Nothing
'
    '----------< 取消データ削除 >------------------------------------------
'    wstr = "Delete * from DCHA010_Gridワーク"
'    wstr = wstr & " Where 取消フラグ=1"
'    GDb.Execute wstr
'
    '----------< テーブル Write >----------------------------------------------
    wstr1 = "Select * from " & wsTbl
    wstr1 = wstr1 & " Where 借入番号 ='" & wsBango & "'"
    Call AdoRecordsetOpen(GDb, wRs1, wstr1)
    If Not wRs1.eof Then

        j = 2 '2回目から始まる
        
        wstr = "Select * from DCHA010_Gridワーク"
        wstr = wstr & " Where テキスト1='" & wsBango & "'"
        wstr = wstr & " Order by 年月日1"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        If Not wRs.eof Then
            Do Until wRs.eof
            
                ws01 = "金利変更" & CStr(j) & "回目年月"
                wRs1(ws01) = P8.FCDate(wRs("年月日1"))
    
                ws01 = "金利" & CStr(j) & "回目"
                wRs1(ws01) = P8.FCDbl(wRs("数値1"))
    
                j = j + 1
                
                wRs.MoveNext
            Loop
            
            If j <= 100 Then
                Do Until j > 100
                    ws01 = "金利変更" & CStr(j) & "回目年月"
                    wRs1(ws01) = Null
        
                    ws01 = "金利" & CStr(j) & "回目"
                    wRs1(ws01) = 0
        
                    j = j + 1
                Loop
            End If
        
            wRs1.Update
        
        Else
        
            If j <= 100 Then
                Do Until j > 100
                    ws01 = "金利変更" & CStr(j) & "回目年月"
                    wRs1(ws01) = Null
        
                    ws01 = "金利" & CStr(j) & "回目"
                    wRs1(ws01) = 0
        
                    j = j + 1
                Loop
            End If
        
            wRs1.Update
        
        End If
        
        wRs.Close
        Set wRs = Nothing
        
    End If
    wRs1.Close
    Set wRs1 = Nothing
    
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    If 新規変更.Caption = "新規" Then
        wslog = "追加"
    ElseIf 新規変更.Caption = "変更" And 削除.Value = 0 Then
        wslog = "更新"
    End If
    GLogStr = "金利変更登録:借入番号=" & P8.FCStr(Me.H_借入番号.Caption) & ","
    GLogStr = GLogStr & "年月日=" & P8.FCStr(Me.L_金利変更年月日) & ","
    GLogStr = GLogStr & "利率=" & P8.FCStr(Me.金利変更利率)
    Call MXA030_LOG_WRITE(pPROGRAM_ID, wslog, GLogStr)
    
'
'
    ' =========================================
    '                 初期設定
    ' =========================================
    If FLG_DEL = True Then
        金利変更年月 = ""
        金利変更利率 = 0
    
        L_金利変更年月日.Caption = ""
    End If
    
    '取消 = 0
    
    'ワークテーブル作成とワークデータセット
    Call 金利ワークテーブル作成
    
    Call 金利変更画面セット
    
    Call AdodcRefresh
        
    Call CEkey.SetFs(金利変更年月, False)
'
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






