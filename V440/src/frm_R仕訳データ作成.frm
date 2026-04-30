VERSION 5.00
Begin VB.Form frm_R仕訳データ作成 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "月次仕訳データ　出力"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   Icon            =   "frm_R仕訳データ作成.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox 終了日 
      Alignment       =   1  '右揃え
      Height          =   330
      IMEMode         =   2  'ｵﾌ
      Left            =   2760
      TabIndex        =   1
      Text            =   "HH年MM月"
      Top             =   1320
      Width           =   1095
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
      TabIndex        =   3
      Top             =   2160
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1920
      Width           =   495
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
      Caption         =   "月次仕訳データ　出力"
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
   Begin VB.Label Label4 
      Caption         =   "～"
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "終了年月To"
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
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label L_実行日 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "開始年月From"
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
      Top             =   960
      Width           =   2535
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
      TabIndex        =   6
      Top             =   1920
      Width           =   2535
   End
End
Attribute VB_Name = "frm_R仕訳データ作成"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "frm_R仕訳データ作成"
'
Dim wRs As ADODB.Recordset
Dim wstr As String

'------------------------------------------------
' Form_Load
'------------------------------------------------
  Private Sub Form_Load()
'
    GRpt.帳票名 = "仕訳表 -月次処理-"
    
    Me.Left = G_LEFT
    Me.Top = G_TOP
    
    実行日 = ""
    終了日 = ""
    Check1 = 0
    
    If G基本情報.日付入力区分 = "0" Then
    '和暦
'        Check2 = 0
    Else
    '西暦
'        Check2 = 1
    End If
'
    実行日 = Format(Now, Gfmt年月)
    終了日 = Format(Now, Gfmt年月)
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
            MsgBox "年月が違います"
            実行日 = "": Call CEkey.SetFs(実行日, True)
            Exit Sub
        End If
    End If
'
    Call CEkey.AllSelect
'
End Sub

'------------------------------------------------
' 終了日_LostFocus
'------------------------------------------------
Private Sub 終了日_LostFocus()
'
    Dim wi01 As Integer, wi02 As Integer
    Dim ws01 As String, wsTuki As String, wsNen As String
'
    If 終了日 = "" Then
        Exit Sub
    End If
'
    If InStrRev(終了日, "年") Then
        GVar1 = C年月日.平成To西暦("", 終了日)
        If GVar1 = 0 Then
            MsgBox "年月を入力してください"
            終了日 = "": Call CEkey.SetFs(終了日, True)
            Exit Sub
        End If
    
        '23年9月→23年09月
        wi01 = Len(終了日)
        wi02 = InStr(終了日, "年")
        ws01 = Mid(終了日, wi02 + 1, wi01 - wi02)
        wsTuki = Right("00" & ws01, 3)
        wsNen = Left(終了日, wi02)
        終了日 = wsNen & wsTuki
    
    Else
        If Len(終了日) < 3 Then
            MsgBox "年月を入力してください"
            終了日 = "": Call CEkey.SetFs(終了日, True)
            Exit Sub
        End If
        
        終了日 = C年月日.FormatDate("年月", 終了日)
        If C年月日.平成To西暦("年月", 終了日) = 0 Then
            MsgBox "年月が違います"
            終了日 = "": Call CEkey.SetFs(終了日, True)
            Exit Sub
        End If
    End If
'
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
    Dim wsRet As String
    
    Dim ws01 As String, ws02 As String
    Dim wsSuii As String, wsSentaku As String, wsSagyo As String, wsJiseki As String
    Dim wsShukei As String, wsSitei As String
    Dim wsUri As String, wsUri2 As String
    Dim wsKar As String, wsKar2 As String
    Dim wsSet As String, wsSet2 As String
    Dim wsKin As String, wskin2 As String
    Dim wsSeR As String, wsSeR2 As String
    Dim wsLea As String, wsLea2 As String
    
    Dim wdate As Date, wDate1 As Date
'
    On Error GoTo 入力チェック_ERR
'
    Call MXA030_GRPTCLEAR
'
    ws01 = P8.FCStr(実行日)
    GVar1 = C年月日.平成To西暦("年月", ws01)

    If ws01 = "" Then
        MsgBox "年月を入力してください"
        実行日 = "": Call CEkey.SetFs(実行日, True)
        Exit Function
    
    End If
    
    ws01 = P8.FCStr(終了日)
    GVar1 = C年月日.平成To西暦("年月", ws01)

    If ws01 = "" Then
        MsgBox "年月を入力してください"
        終了日 = "": Call CEkey.SetFs(終了日, True)
        Exit Function
    
    End If
'
    '同一年度CHECK
    wdate = C年月日.平成To西暦("年月日", P8.FCStr(実行日))
    'ws01 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "e")
    '2019/01/15 日付入力区分 仕様変更
    If G基本情報.日付入力区分 = "0" Then
    '和暦
        ws01 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "e")
    Else
    '西暦
        ws01 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "y")
    End If
    
    wdate = C年月日.平成To西暦("年月日", P8.FCStr(終了日))
    'ws02 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "e")
    If G基本情報.日付入力区分 = "0" Then
    '和暦
        ws02 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "e")
    Else
    '西暦
        ws02 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "y")
    End If
    
    If ws01 <> ws02 Then
        'MsgBox "同一年度内の年月を入力してください"
        '終了日 = "": Call CEkey.SetFs(終了日, True)
        'Exit Function
    
        GRet = MsgBox("指定年月が同一年度内ではありません" & vbCrLf & vbCrLf & "出力しますか？", vbYesNo + vbQuestion)
        If GRet = vbNo Then
            終了日 = "": Call CEkey.SetFs(終了日, True)
            Exit Function
        End If
    End If
'
    '終了年月 CHECK
'    If InStrRev(終了日, "年") Then
'        GVar1 = C年月日.平成To西暦("", 終了日)
'        wdate = CDate(GVar1)
'    Else
'        wdate = CDate(終了日 & "/01")
'    End If
'    'wDate1 = DateAdd("m", 1, wdate) '翌月
'    'wdate = DateAdd("d", -1, wDate1)
'
'    wdate = MBA010_締日年月日(wdate)
'    wDate1 = DateAdd("d", 1, wdate)
'    wDate1 = MBA010_対象年月(wDate1) '翌月
'
'    Call C休日.計算(wdate, 0)
'    If C休日.休日 = True Then
'        If InStrRev(GRpt.テキスト_02, "年") Then
'            wsRet = "月末日「" & Format(wdate, "ee年mm月dd日") & "」は金融機関の休日になります"
'        Else
'            wsRet = "月末日「" & wdate & "」は金融機関の休日になります"
'        End If
'
'        GRet = MsgBox(wsRet & vbCrLf & vbCrLf & "終了月を翌月にセットしますか？", vbYesNo + vbQuestion)
'        If GRet = vbYes Then
'            If InStrRev(終了日, "年") Then
'                終了日 = Format(wDate1, "ee年mm月")
'            Else
'                終了日 = Format(wDate1, "yyyy/mm")
'            End If
'        End If
'    End If
'
    'If FLG_Check = True Then
    '    MsgBox "指定された内容に誤りがあります"
    '
    '    Exit Function
    'End If
'
    FLG_Check = False
'
    ' =========================================
    '                    後処理
    ' =========================================
    GRpt.テキスト_01 = 実行日
    GRpt.テキスト_02 = 終了日
    GRpt.CSV = Check1
'    GRpt.チェック_02 = Check2
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
    
    Dim wdate As Date
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
    If InStrRev(実行日, "年") Then
        GVar1 = C年月日.平成To西暦("", 実行日)
        If GVar1 = 0 Then
            MsgBox "年月を入力してください"
            実行日 = "": Call CEkey.SetFs(実行日, True)
            Exit Sub
        End If
    Else
        If Len(実行日) < 3 Then
            MsgBox "年月を入力してください"
            実行日 = "": Call CEkey.SetFs(実行日, True)
            Exit Sub
        End If
        
        実行日 = C年月日.FormatDate("年月", 実行日)
        If C年月日.平成To西暦("年月", 実行日) = 0 Then
            MsgBox "年月が違います"
            実行日 = "": Call CEkey.SetFs(実行日, True)
            Exit Sub
        End If
    End If
    
    If InStrRev(終了日, "年") Then
        GVar1 = C年月日.平成To西暦("", 終了日)
        If GVar1 = 0 Then
            MsgBox "年月を入力してください"
            終了日 = "": Call CEkey.SetFs(終了日, True)
            Exit Sub
        End If
    ElseIf 終了日 = "" Then
        終了日 = 実行日
    Else
        If Len(終了日) < 3 Then
            MsgBox "年月を入力してください"
            終了日 = "": Call CEkey.SetFs(終了日, True)
            Exit Sub
        End If
        
        終了日 = C年月日.FormatDate("年月", 終了日)
        If C年月日.平成To西暦("年月", 終了日) = 0 Then
            MsgBox "年月が違います"
            終了日 = "": Call CEkey.SetFs(終了日, True)
            Exit Sub
        End If
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
    Set rpt = New RDH010_仕訳データ
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
    GLogStr = ""
    GLogStr = GLogStr & "開始年月=" & GRpt.テキスト_01 & ","
    GLogStr = GLogStr & "終了年月=" & GRpt.テキスト_02 & ","
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
' 閉じる_Click
'------------------------------------------------
Private Sub 閉じる_Click()
'
    Call MXA030_MCLEAR
'
    Unload Me
'
End Sub

