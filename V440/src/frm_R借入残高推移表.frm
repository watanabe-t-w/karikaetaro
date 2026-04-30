VERSION 5.00
Begin VB.Form frm_R借入残高推移表 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "借入残高推移表　出力"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10185
   Icon            =   "frm_R借入残高推移表.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton 帳票出力設定 
      Caption         =   "帳票出力設定"
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
      Left            =   8040
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   240
      Width           =   1815
   End
   Begin VB.ComboBox 借入金種別名 
      Height          =   300
      IMEMode         =   1  'ｵﾝ
      Left            =   2760
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   4
      Top             =   2640
      Width           =   4215
   End
   Begin VB.ComboBox 部門名 
      Height          =   300
      IMEMode         =   1  'ｵﾝ
      Left            =   2760
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   5
      Top             =   3000
      Width           =   4215
   End
   Begin VB.ComboBox 金融機関名 
      Height          =   300
      IMEMode         =   1  'ｵﾝ
      Left            =   2760
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   6
      Top             =   3360
      Width           =   4215
   End
   Begin VB.CheckBox Check6 
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   6120
      Width           =   495
   End
   Begin VB.CheckBox Check5 
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   5640
      Width           =   495
   End
   Begin VB.CheckBox Check2 
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   4560
      Width           =   495
   End
   Begin VB.ComboBox 銀行名 
      Height          =   300
      IMEMode         =   1  'ｵﾝ
      Left            =   2760
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   7
      Top             =   3720
      Width           =   4215
   End
   Begin VB.CheckBox Check4 
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   5280
      Width           =   495
   End
   Begin VB.ComboBox 集計区分 
      Height          =   300
      IMEMode         =   1  'ｵﾝ
      Left            =   2760
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   3
      Top             =   2280
      Width           =   3135
   End
   Begin VB.ComboBox 金融リストラ番号 
      Height          =   300
      IMEMode         =   1  'ｵﾝ
      Left            =   2760
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   1920
      Width           =   3135
   End
   Begin VB.ComboBox 出力表種別 
      Height          =   300
      IMEMode         =   1  'ｵﾝ
      Left            =   2760
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   960
      Width           =   3135
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox 実行日 
      Alignment       =   1  '右揃え
      Height          =   330
      IMEMode         =   2  'ｵﾌ
      Left            =   2760
      TabIndex        =   1
      Text            =   "HH年"
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton 実行 
      Caption         =   "実行（F11)"
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
      Left            =   6120
      TabIndex        =   14
      Top             =   6000
      Width           =   1815
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
      Left            =   8040
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CheckBox Check3 
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   4920
      Width           =   495
   End
   Begin 借換たろう.ZU050_Button ZU050_Button1 
      Height          =   615
      Left            =   120
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "借入残高推移表　出力"
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
   Begin VB.Label Label11 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "借入金種別名"
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
      TabIndex        =   30
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label10 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "部門名"
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
      TabIndex        =   29
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label15 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "金融機関名"
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
      TabIndex        =   28
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label9 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "(社債)手数料､保証料出力"
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
      Left            =   240
      TabIndex        =   27
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "決算用で出力"
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
      TabIndex        =   26
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label Label8 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "千円単位で出力"
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
      TabIndex        =   25
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label45 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "銀行名"
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
      TabIndex        =   24
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label L_金利SM 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  '実線
      Caption         =   "金利ｼﾐｭﾚｰｼｮﾝ"
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
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "集計区分"
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
      TabIndex        =   22
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "ｼﾐｭﾚｰｼｮﾝ番号"
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
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label7 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "出力表種別"
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
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "推移開始年度"
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
      TabIndex        =   19
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "個別借入金を表示"
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
      TabIndex        =   18
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "CSV出力"
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
      TabIndex        =   17
      Top             =   5280
      Width           =   2535
   End
End
Attribute VB_Name = "frm_R借入残高推移表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "frm_R借入残高推移表"
'
Dim wRs As ADODB.Recordset
Dim wstr As String
'
'------------------------------------------------
' Form_Load
'------------------------------------------------
  Private Sub Form_Load()
'
    Dim wi01 As Integer
'
    GRpt.帳票名 = "借入残高推移表"
    
    Me.Left = G_LEFT
    Me.Top = G_TOP
    
    実行日 = ""
    出力表種別.Clear
    金融リストラ番号.Clear
    集計区分.Clear
    借入金種別名.Clear
    部門名.Clear
    金融機関名.Clear
    銀行名.Clear
    
    Check1 = 0
    If G基本情報.借入金管理区分 = XMXA020_区分("借入金管理区分", "決算用") Then
        Check1 = 1
    End If
    
    Check2 = 0
    Check3 = 0
    Check4 = 0
    Check5 = 0
    Check6 = 0

    Call MFA001_出力表種別(出力表種別)
    Call MFA001_金融リストラ番号(金融リストラ番号)
    Call MFA001_集計区分(集計区分)
    Call MFA001_借入金種別名(借入金種別名)
    Call MFA001_部門名(部門名)
    Call MFA001_金融機関名(金融機関名)
    Call MFA001_銀行区分(銀行名)
'
    L_金利SM.Visible = False
    Check6.Visible = False
    
    '金利GR
    If GStr = "金利GR" Then
        L_金利SM.Visible = True
        Check6 = 1
        Check6.Visible = True
    End If
'
    '実行日 = Replace(Format(C年月日.年度開始年月日(C年月日.年度変換(CStr(Now)), "西暦"), Gfmt年), "年", "")
    wi01 = C年月日.年度変換(Format(Now, "yyyy/mm/dd"))
    If wi01 <> 0 Then
        実行日 = Replace(Format(CDate(CStr(wi01) & "/01/01"), Gfmt年), "年", "")
    End If
'
End Sub

'------------------------------------------------
' Form_Activate
'------------------------------------------------
Private Sub Form_Activate()
'
    '*** 画面のちらつきをなくす為の Doevents
    DoEvents
    
    Call CEkey.AllSelect
End Sub

'------------------------------------------------
' Form_KeyDown
'------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
    If KeyCode = vbKeyF11 Then
        Call 実行_Click
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

Private Sub 実行日_LostFocus()
    If G基本情報.日付入力区分 = "0" Then
    '和暦
        Call P8.FCControlLeft(実行日, 2)
    Else
    '西暦
        Call P8.FCControlLeft(実行日, 4)
    End If
End Sub

'------------------------------------------------
' 入力チェック
'------------------------------------------------
Private Function 入力チェック() As Boolean
'
    Dim j As Integer
    Dim FLG_Check As Boolean
    
    Dim ws01 As String, ws02 As String
    Dim wsSuii As String, wsSentaku As String, wsSagyo As String, wsJiseki As String
    Dim wsShukei As String, wsSitei As String
    Dim wsShube As String, wsBumon As String, wsKinyu As String, wsGinko As String
    Dim wsUri As String, wsUri2 As String
    Dim wsKar As String, wsKar2 As String
    Dim wsSet As String, wsSet2 As String
    Dim wsKin As String, wskin2 As String
    Dim wsSeR As String, wsSeR2 As String
    Dim wsLea As String, wsLea2 As String
    
    '金融機関番名と銀行名
    Dim Kinyusu As Integer   '金融機関名の文字数
    Dim Ginkosu As Integer   '銀行名の文字数
    Dim Str As String       '検索対象文字列
    Dim Striti As Integer   '文字列位置
'
    On Error GoTo 入力チェック_ERR
'
    入力チェック = False
    FLG_Check = False

    Call MXA030_GRPTCLEAR
'
    ws01 = P8.FCStr(実行日)
    wsSuii = P8.FCStr(出力表種別.Text)
    wsKin = P8.FCStr(金融リストラ番号.Text)
    wsSitei = P8.FCStr(銀行名.Text)
    wsShukei = P8.FCStr(集計区分.Text)
    
    wsShube = P8.FCStr(借入金種別名.Text)
    wsBumon = P8.FCStr(部門名.Text)
    wsKinyu = P8.FCStr(金融機関名.Text)
    wsGinko = P8.FCStr(銀行名.Text)
    
    If Not P8.FIsInt(実行日) Then
        実行日 = ""
        FLG_Check = True:   Call CEkey.SetFs(実行日, True)
    End If
    
    If P8.FCDbl(実行日) <= 0 Then
        実行日 = ""
        FLG_Check = True:   Call CEkey.SetFs(実行日, True)
    End If
        
    'コンボ
    Kinyusu = Len(金融機関名)
    Ginkosu = Len(銀行名)
    
    If Kinyusu > 0 And Ginkosu > 0 Then
        
        Str = 銀行名
        Striti = InStr(Str, " ")
        Striti = Striti - 1
        Str = Left(Str, Striti)
    
        If 金融機関名 <> Str Then
            FLG_Check = True:
        End If
    
    End If
    
    If FLG_Check = True Then
        MsgBox "指定された内容に誤りがあります"

        Exit Function
    End If
'
    FLG_Check = False
'
    ' =========================================
    '                    後処理
    ' =========================================
    GRpt.テキスト_01 = 実行日
    GRpt.推移 = wsSuii
    GRpt.集計 = wsShukei
    GRpt.指定 = wsSitei
    GRpt.金融 = wsKin
    GRpt.借入金管理区分 = Check1
    GRpt.詳細表示 = Check2
    GRpt.千円単位 = Check3
    GRpt.CSV = Check4
    If Check6 = 1 Then
        G金利SM = True
    End If
    
     'コンボ指定
    GRpt.C_種別 = wsShube
    GRpt.C_部門 = wsBumon
    GRpt.C_金融 = wsKinyu
    GRpt.C_銀行 = wsGinko

    Call MFA001_GRrptBunrui("frm_R借入残高推移表")
'
    入力チェック = True
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
入力チェック_ERR:
    pERR_MES = pPROGRAM_ID + "/ 入力チェック() でエラー" + vbCrLf + vbCrLf + _
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
' 実行_Click
'------------------------------------------------
Private Sub 実行_Click()
'
    Dim rpt As Object
    Dim wsRet As String, wsBunrui(3) As String
    Dim j As Integer
'
    On Error GoTo 実行_Click_ERR
'
    ' =========================================
    '           　 入力チェック
    ' =========================================
    If 入力チェック <> True Then
        Exit Sub
    End If
'
    ' =========================================
    '           基本情報ファイル Read
    ' =========================================
    MAA010_基本情報ファイル_Read
    
    G基本情報.借入金管理区分 = GRpt.借入金管理区分
'
    ' =========================================
    '           　 ボタン制御
    ' =========================================
    実行.Enabled = False
    閉じる.Enabled = False
'
    ' =========================================
    '           　 NULLデータ置換
    ' =========================================
    Call MXA030_null置換
'
    Call MXA030_印字テーブルクリア
    Call MXA030_MCLEAR
    GSstrt帳票Msg = ""
'
    ' =========================================
    '              レポート表示
    ' =========================================
    If Check5 = 0 Then
        GRpt.帳票名 = "借入残高推移表"
        If GRpt.推移 = "年次" Or GRpt.推移 = "半期" Then
            'Set rpt = New RDB010_借入残高推移表
            Set rpt = New RDB030_借入残高推移表                     '長短振替
        ElseIf GRpt.推移 = "四半期" Or GRpt.推移 = "月次" Then
            'Set rpt = New RDB010_借入残高推移表
            Set rpt = New RDB040_借入残高推移表                     '長短振替
        End If
    ElseIf Check5 = 1 Then
        GRpt.帳票名 = "社債残高推移表"
        If GRpt.推移 = "年次" Or GRpt.推移 = "半期" Then
            Set rpt = New RDB010_社債残高推移表
        ElseIf GRpt.推移 = "四半期" Or GRpt.推移 = "月次" Then
            Set rpt = New RDB020_社債残高推移表
        End If
    End If
'
    rpt.Show vbModal
'
    Set rpt = Nothing
'
    実行.Enabled = True
    閉じる.Enabled = True
'
    ' =========================================
    '        レポートエラーMsg
    ' =========================================
    If GSstrt帳票Msg <> "" Then
        MsgBox GSstrt帳票Msg, vbInformation
    End If
'
    ' =========================================
    '           基本情報ファイル Read
    ' =========================================
    MAA010_基本情報ファイル_Read
'

'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    For j = 0 To 3
        If GRpt.S_種別 = "分類" & CStr(j + 1) Then
            wsBunrui(j) = "借入金種別"
        ElseIf GRpt.S_部門 = "分類" & CStr(j + 1) Then
            wsBunrui(j) = "部門"
        ElseIf GRpt.S_金融 = "分類" & CStr(j + 1) Then
            wsBunrui(j) = "金融機関"
        ElseIf GRpt.S_銀行 = "分類" & CStr(j + 1) Then
            wsBunrui(j) = "銀行"
        Else
            wsBunrui(j) = "表示しない"
        End If
    Next j
    '
    If GRpt.借入金管理区分 = P8.FCDbl(XMXA020_区分("借入金管理区分", "決算用")) Then
        GLogStr = GLogStr & "帳票指示=決算用,"
    Else
        GLogStr = GLogStr & "帳票指示=管理用,"
    End If
    GLogStr = GLogStr & "出力表種別=" & GRpt.推移 & ","
    GLogStr = GLogStr & "推移開始年度=" & GRpt.テキスト_01 & ","
    GLogStr = GLogStr & "ｼﾐｭﾚｰｼｮﾝ番号=" & GRpt.金融 & ","
    GLogStr = GLogStr & "集計区分=" & GRpt.指定 & ","
    GLogStr = GLogStr & "借入金種別名=" & GRpt.C_種別 & ","
    GLogStr = GLogStr & "部門名=" & GRpt.C_部門 & ","
    GLogStr = GLogStr & "金融機関名=" & GRpt.C_金融 & ","
    GLogStr = GLogStr & "銀行名=" & GRpt.C_銀行 & ","
    GLogStr = GLogStr & "個別借入金表示=" & GRpt.詳細表示 & ","
    GLogStr = GLogStr & "千円単位=" & GRpt.千円単位 & ","
    GLogStr = GLogStr & "CSV出力=" & GRpt.CSV & ","
    GLogStr = GLogStr & "金利ｼﾐｭﾚｰｼｮﾝ=" & Check6.Value & ","
    GLogStr = GLogStr & "分類1=" & wsBunrui(0) & ","
    GLogStr = GLogStr & "分類2=" & wsBunrui(1) & ","
    GLogStr = GLogStr & "分類3=" & wsBunrui(2) & ","
    GLogStr = GLogStr & "分類4=" & wsBunrui(3)
    Call MXA030_LOG_WRITE(GRpt.帳票名, "帳票", GLogStr)
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
実行_Click_ERR:
    pERR_MES = pPROGRAM_ID + "/ 実行_Click() でエラー" + vbCrLf + vbCrLf + _
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
' 帳票出力設定_Click
'------------------------------------------------
Private Sub 帳票出力設定_Click()
'
    Me.Enabled = False
    Set GForm = frm_R借入残高推移表
    frm_R帳票出力設定.Show
'
End Sub

'------------------------------------------------
' 閉じる_Click
'------------------------------------------------
Private Sub 閉じる_Click()
'
    Call MXA030_MCLEAR
'
    Unload Me
'
End Sub


