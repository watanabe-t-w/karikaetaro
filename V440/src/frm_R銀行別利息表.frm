VERSION 5.00
Begin VB.Form frm_R銀行別利息表 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "銀行別利息表　出力"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   Icon            =   "frm_R銀行別利息表.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   240
      Width           =   1815
   End
   Begin VB.ComboBox 借入金種別名 
      Height          =   300
      IMEMode         =   1  'ｵﾝ
      Left            =   2760
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   1800
      Width           =   4215
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2280
      Width           =   495
   End
   Begin VB.ComboBox 実行月 
      Height          =   300
      Left            =   1560
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox 実行日 
      Alignment       =   1  '右揃え
      Height          =   330
      IMEMode         =   2  'ｵﾌ
      Left            =   1560
      TabIndex        =   0
      Text            =   "HH年"
      Top             =   960
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
      TabIndex        =   5
      Top             =   2760
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   2640
      Width           =   495
   End
   Begin 借換たろう.ZU050_Button ZU050_Button1 
      Height          =   615
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "銀行別利息表　出力"
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
   Begin VB.Label Label2 
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
      TabIndex        =   12
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label3 
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
      TabIndex        =   11
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "月"
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
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "年度"
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
      TabIndex        =   9
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label9 
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
      TabIndex        =   8
      Top             =   2640
      Width           =   2535
   End
End
Attribute VB_Name = "frm_R銀行別利息表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'杉本倉庫仕様
Private Const pPROGRAM_ID As String = "frm_R銀行別利息表"
'
Dim wRs As ADODB.Recordset
Dim wstr As String

'------------------------------------------------
' Form_Load
'------------------------------------------------
  Private Sub Form_Load()
'
    Dim wi01 As Integer
'
    GRpt.帳票名 = "銀行別利息表"
    
    Me.Left = G_LEFT
    Me.Top = G_TOP
    
    実行日 = ""
    借入金種別名.Clear
    Check1 = 0
    Check2 = 0
'
    wi01 = C年月日.年度変換(Format(Now, "yyyy/mm/dd"))
    If wi01 <> 0 Then
        実行日 = Replace(Format(CDate(CStr(wi01) & "/01/01"), Gfmt年), "年", "")
    End If
'
    Call 実行月_セット
    Call MFA001_借入金種別名(借入金種別名)
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

'------------------------------------------------
' 実行日_LostFocus
'------------------------------------------------
Private Sub 実行日_LostFocus()
    If G基本情報.日付入力区分 = "0" Then
    '和暦
        Call P8.FCControlLeft(実行日, 2)
    Else
    '西暦
        Call P8.FCControlLeft(実行日, 4)
    End If
'
    Call CEkey.AllSelect
'
End Sub

'------------------------------------------------
' 実行月_セット
'------------------------------------------------
Private Sub 実行月_セット()
'
    Dim j As Integer
    Dim wi01 As Integer, wiCnt As Integer
    Dim w開始年月日 As Date, wdate As Date
'
    w開始年月日 = C年月日.年度開始年月日(Format(Now, "yyyy"), "西暦")
    
    実行月.Clear
    
    If G基本情報.決算サイクル = 1 _
    Or G基本情報.決算サイクル = 3 _
    Or G基本情報.決算サイクル = 6 Then
        
        wdate = DateAdd("m", -1, w開始年月日)
        wiCnt = 12 / G基本情報.決算サイクル
        
        For j = 1 To wiCnt
            wdate = DateAdd("m", G基本情報.決算サイクル, wdate)
            wi01 = CInt(Format(wdate, "mm"))
            
            実行月.AddItem wi01
        Next
        
    Else
    '年次 G基本情報.決算サイクル=12
        実行月.AddItem G基本情報.決算月
    End If
    
    実行月 = 実行月.List(0)
'
End Sub

'------------------------------------------------
' 入力チェック
'------------------------------------------------
Private Function 入力チェック() As Boolean
'
    Dim j As Integer
    Dim FLG_Check As Boolean
    
    Dim wdate As Date
    Dim wi01 As Integer
    Dim ws01 As String, ws02 As String
    Dim wsShube As String
'
    On Error GoTo 入力チェック_ERR
'
    Call MXA030_GRPTCLEAR
'
    If Not P8.FIsInt(実行日) Then
        実行日 = ""
        FLG_Check = True:   Call CEkey.SetFs(実行日, True)
    End If
    
    If P8.FCDbl(実行日) <= 0 Then
        実行日 = ""
        FLG_Check = True:   Call CEkey.SetFs(実行日, True)
    End If
    
    ws01 = ""
    wi01 = 実行月.ListIndex
    If G基本情報.日付入力区分 = "0" Then
    '和暦
        wdate = DateAdd("m", -1, C年月日.年度開始年月日(実行日, "平成"))
        ws01 = Format(DateAdd("m", 3 * (wi01 + 1), wdate), "ee年mm月")
    Else
    '西暦
        wdate = DateAdd("m", -1, C年月日.年度開始年月日(実行日, "西暦"))
        ws01 = Format(DateAdd("m", 3 * (wi01 + 1), wdate), "yyyy/mm")
    End If
    
    If InStrRev(ws01, "年") Then
        GVar1 = C年月日.平成To西暦("", ws01)
        If GVar1 = 0 Then
            MsgBox "年月を入力してください"
            FLG_Check = True: Call CEkey.SetFs(実行日, True)
        End If
    
'        ws02 = DateAdd("m", -5, CDate(GVar1))
'        ws02 = Format(ws02, "ee年mm月")
    Else
        If Len(ws01) < 3 Then
            MsgBox "年月を入力してください"
            FLG_Check = True: Call CEkey.SetFs(実行日, True)
        End If
        
        ws01 = C年月日.FormatDate("年月", ws01)
        If C年月日.平成To西暦("年月", ws01) = 0 Then
            MsgBox "年月が違います"
            FLG_Check = True: Call CEkey.SetFs(実行日, True)
        End If
        
'        ws02 = DateAdd("m", -5, CDate(ws01 & "/01"))
'        ws02 = Format(ws02, "yyyy/mm")
    End If
    
    wsShube = P8.FCStr(借入金種別名.Text)
    
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
    GRpt.テキスト_01 = P8.FCDbl(実行日)
    GRpt.テキスト_02 = ws01
    
    'コンボ指定
    GRpt.C_種別 = wsShube
    GRpt.推移 = "四半期"
    
    GRpt.詳細表示 = Check1
    GRpt.CSV = Check2
    
    Call MFA001_GRrptBunrui("frm_R銀行別利息表")
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
    Dim wsRet As String

    Dim ws管理区分 As String
'
    On Error GoTo 実行_Click_ERR
'
    ' =========================================
    '           基本情報ファイル Read
    ' =========================================
    MAA010_基本情報ファイル_Read
'
    If 実行日 = "" Then
        Exit Sub
    End If
'
    ' =========================================
    '           　 入力チェック
    ' =========================================
    If 入力チェック <> True Then
        Exit Sub
    End If
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
    'G基本情報.借入金管理区分=決算用で出力
    ws管理区分 = G基本情報.借入金管理区分
    G基本情報.借入金管理区分 = XMXA020_区分("借入金管理区分", "決算用")
    
    Set rpt = New RDH030_銀行別利息表
'
    rpt.Show vbModal
'
    Set rpt = Nothing
'
    'G基本情報.借入金管理区分 元に戻す
    G基本情報.借入金管理区分 = ws管理区分
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
    '               LOG_WRITE
    ' =========================================
    GLogStr = ""
    GLogStr = GLogStr & "開始年月=" & GRpt.テキスト_01 & ","
    GLogStr = GLogStr & "CSV出力=" & GRpt.CSV & ","
    Call MXA030_LOG_WRITE(GRpt.帳票名, "帳票", GLogStr)

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
    Set GForm = frm_R銀行別利息表
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

