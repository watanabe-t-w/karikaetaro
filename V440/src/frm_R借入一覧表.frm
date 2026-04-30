VERSION 5.00
Begin VB.Form frm_R借入一覧表 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "借入一覧表　出力"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10185
   Icon            =   "frm_R借入一覧表.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3240
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton 全件 
      Caption         =   "登録データ一覧表示"
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
      Left            =   4080
      TabIndex        =   13
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CheckBox Check2 
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   2280
      Width           =   495
   End
   Begin VB.ComboBox 金融リストラ番号 
      Height          =   300
      IMEMode         =   1  'ｵﾝ
      Left            =   2760
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   1440
      Width           =   3135
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox 実行日 
      Alignment       =   1  '右揃え
      Height          =   330
      IMEMode         =   2  'ｵﾌ
      Left            =   2760
      TabIndex        =   0
      Text            =   "HH年MM月"
      Top             =   960
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   2640
      Width           =   495
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1815
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
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin 借換たろう.ZU050_Button ZU050_Button1 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "借入一覧表　出力"
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
   Begin VB.Label Label1 
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
      TabIndex        =   12
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label45 
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
      TabIndex        =   10
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "年月"
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
      Top             =   960
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
      TabIndex        =   7
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label2 
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
      TabIndex        =   6
      Top             =   1920
      Width           =   2535
   End
End
Attribute VB_Name = "frm_R借入一覧表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "frm_R借入一覧表"
'
Dim wRs As ADODB.Recordset
Dim wstr As String

'------------------------------------------------
' Form_Load
'------------------------------------------------
  Private Sub Form_Load()
'
    GRpt.帳票名 = "借入一覧表"
    
    Me.Left = G_LEFT
    Me.Top = G_TOP
    
    実行日 = Format(Now, Gfmt年月)
    金融リストラ番号.Clear
    
    Check1 = 0
    If G基本情報.借入金管理区分 = XMXA020_区分("借入金管理区分", "決算用") Then
        Check1 = 1
    End If
    
    Check2 = 0
    Check3 = 0

    Call MFA001_金融リストラ番号(金融リストラ番号)
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
'
    Dim wi01 As Integer, wi02 As Integer
    Dim ws01 As String, wsTuki As String, wsNen As String
'
    If 実行日 = "" Then
        Exit Sub
    End If
'
    If InStrRev(実行日, "年") Then
        GVar1 = C年月日.平成To西暦("", 実行日)
        If GVar1 = 0 Then
            MsgBox "年月を入力してください"
            実行日 = "": Call CEkey.SetFs(実行日, True)
            Exit Sub
        End If
    
        '23年9月→23年09月
        wi01 = Len(実行日)
        wi02 = InStr(実行日, "年")
        ws01 = Mid(実行日, wi02 + 1, wi01 - wi02)
        wsTuki = Right("00" & ws01, 3)
        wsNen = Left(実行日, wi02)
        実行日 = wsNen & wsTuki
    
    Else
        If Len(実行日) < 3 Then
            MsgBox "年月を入力してください"
            実行日 = "": Call CEkey.SetFs(実行日, True)
            Exit Sub
        End If
        
        実行日 = C年月日.FormatDate("年月", 実行日)
        If C年月日.平成To西暦("年月", 実行日) = 0 Then
            MsgBox "月日が違います"
            実行日 = "": Call CEkey.SetFs(実行日, True)
            Exit Sub
        End If
    End If
    
    Call CEkey.AllSelect
'
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
    Dim wsUri As String, wsUri2 As String
    Dim wsKar As String, wsKar2 As String
    Dim wsSet As String, wsSet2 As String
    Dim wsKin As String, wskin2 As String
    Dim wsSeR As String, wsSeR2 As String
    Dim wsLea As String, wsLea2 As String
'
    On Error GoTo 入力チェック_ERR
'
    入力チェック = False
    FLG_Check = False
    
    Call MXA030_GRPTCLEAR
'
    wsKin = P8.FCStr(金融リストラ番号.Text)
    
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
    GRpt.金融 = wsKin
    GRpt.借入金管理区分 = Check1
    GRpt.千円単位 = Check2
    GRpt.CSV = Check3

    Call MFA001_GRrptBunrui("frm_R借入一覧表")
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
    Dim wsRet As String, wsBunrui(1) As String
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
    GRpt.帳票名 = "借入一覧表"
    Set rpt = New RDA020_借入一覧表
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
    '               LOG_WRITE
    ' =========================================
    For j = 0 To 1
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
    GLogStr = ""
    GLogStr = "年月=" & GRpt.テキスト_01 & ","
    GLogStr = GLogStr & "ｼﾐｭﾚｰｼｮﾝ番号=" & GRpt.金融 & ","
    GLogStr = GLogStr & "分類1=" & wsBunrui(0) & ","
    GLogStr = GLogStr & "分類2=" & wsBunrui(1)
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
' 全件_Click
'------------------------------------------------
Private Sub 全件_Click()
'
    Dim rpt As Object
    Dim wsRet As String, wsBunrui(1) As String
    Dim j As Integer
'
    On Error GoTo 全件_Click_ERR
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
    '              レポート表示
    ' =========================================
    GRpt.帳票名 = "借入一覧表(全件)"
    Set rpt = New RDA020_借入一覧表_全件
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
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    For j = 0 To 1
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
    GLogStr = ""
    If GRpt.借入金管理区分 = P8.FCDbl(XMXA020_区分("借入金管理区分", "決算用")) Then
        GLogStr = GLogStr & "帳票指示=決算用,"
    Else
        GLogStr = GLogStr & "帳票指示=管理用,"
    End If
    GLogStr = GLogStr & "ｼﾐｭﾚｰｼｮﾝ番号=" & GRpt.金融 & ","
    GLogStr = GLogStr & "分類1=" & wsBunrui(0) & ","
    GLogStr = GLogStr & "分類2=" & wsBunrui(1)
    Call MXA030_LOG_WRITE(GRpt.帳票名, "帳票", GLogStr)
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
全件_Click_ERR:
    pERR_MES = pPROGRAM_ID + "/ 全件_Click() でエラー" + vbCrLf + vbCrLf + _
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
    Set GForm = frm_R借入一覧表
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




