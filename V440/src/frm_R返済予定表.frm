VERSION 5.00
Begin VB.Form frm_R返済予定表 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "返済予定表　出力"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10185
   Icon            =   "frm_R返済予定表.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check2 
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      Top             =   3000
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Left            =   2880
      TabIndex        =   14
      Top             =   2640
      Width           =   495
   End
   Begin VB.ComboBox 集計区分 
      Height          =   300
      IMEMode         =   1  'ｵﾝ
      Left            =   2760
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   3
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox 実行日 
      Height          =   330
      IMEMode         =   2  'ｵﾌ
      Left            =   2760
      TabIndex        =   0
      Text            =   "HH年MM月DD日"
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox 実行日2 
      Height          =   330
      IMEMode         =   2  'ｵﾌ
      Left            =   2760
      TabIndex        =   1
      Text            =   "HH年MM月DD日"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ComboBox 銀行名 
      Height          =   300
      IMEMode         =   1  'ｵﾝ
      Left            =   2760
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   1800
      Width           =   4215
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
      TabIndex        =   4
      Top             =   3240
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1815
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
      Caption         =   "返済予定表　出力"
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
   Begin VB.Label Label8 
      Caption         =   "返済年月日(HHMMDD)"
      Height          =   255
      Left            =   5280
      TabIndex        =   16
      Top             =   1080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "返済年月日From"
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
      TabIndex        =   13
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "返済年月日To"
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
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "～"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   1080
      Width           =   255
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
      TabIndex        =   10
      Top             =   2640
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
      TabIndex        =   9
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "集計"
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
      Top             =   2160
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
      Top             =   3000
      Width           =   2535
   End
End
Attribute VB_Name = "frm_R返済予定表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "frm_R返済予定表"
'
Dim wRs As ADODB.Recordset
Dim wstr As String

'------------------------------------------------
' Form_Load
'------------------------------------------------
  Private Sub Form_Load()
'
    GRpt.帳票名 = "借入金返済予定表"
    
    Me.Left = G_LEFT
    Me.Top = G_TOP
    
    実行日 = ""
    実行日2 = ""
    集計区分.Clear
    銀行名.Clear
    Check1 = 0
    Check2 = 0

    Call MFA001_集計区分2(集計区分)
    Call MFA001_銀行区分(銀行名)
'
    実行日 = Format(Format(Now, "yyyy/mm/01"), Gfmt年月日)
    実行日2 = Format(DateAdd("d", -1, DateAdd("m", 1, DateValue(Format(Now, "yyyy/mm/01")))), Gfmt年月日)

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
    Dim ws01 As String
'
    If 実行日 = "" Then
        Exit Sub
    End If
'
    If InStrRev(実行日, "年") Then
        GVar1 = C年月日.平成To西暦("", 実行日)
        If GVar1 = 0 Then
            MsgBox "返済年月日Fromを入力してください"
            実行日 = "": Call CEkey.SetFs(実行日, True)
            Exit Sub
        End If
    Else
        If Len(実行日) < 4 Then
            MsgBox "返済年月日Fromを入力してください"
            実行日 = "": Call CEkey.SetFs(実行日, True)
            Exit Sub
        End If
        
        If Len(実行日) <= 4 Then
            ws01 = Right$("0000" & 実行日, 2)
            If ws01 < "01" Or ws01 > "12" Then
                MsgBox "返済年月日Fromを入力してください"
                実行日 = "": Call CEkey.SetFs(実行日, True)
                Exit Sub
            End If
        
            実行日 = 実行日 & "01"
            
        ElseIf Len(実行日) > 4 Then
            ws01 = Right$("000000" & 実行日, 2)
            If ws01 < "01" Or ws01 > "31" Then
                MsgBox "返済年月日Fromを入力してください"
                実行日 = "": Call CEkey.SetFs(実行日, True)
                Exit Sub
            End If
        End If
        
        実行日 = C年月日.FormatDate("年月日", 実行日)
        If C年月日.平成To西暦("年月日", 実行日) = 0 Then
            MsgBox "返済年月日Fromが違います"
            実行日 = "": Call CEkey.SetFs(実行日, True)
            Exit Sub
        End If
    End If
'
    Call CEkey.AllSelect
'
End Sub

Private Sub 実行日2_LostFocus()
'
    Dim ws01 As String
'
    If 実行日2 = "" Then
        Exit Sub
    End If
'
    If InStrRev(実行日2, "年") Then
        GVar1 = C年月日.平成To西暦("", 実行日2)
        If GVar1 = 0 Then
            MsgBox "返済年月日Toを入力してください"
            実行日2 = "": Call CEkey.SetFs(実行日2, True)
            Exit Sub
        End If
    Else
        If Len(実行日2) < 4 Then
            MsgBox "返済年月日Toを入力してください"
            実行日2 = "": Call CEkey.SetFs(実行日2, True)
            Exit Sub
        End If
        
        If Len(実行日2) <= 4 Then
            ws01 = Right$("0000" & 実行日2, 2)
            If ws01 < "01" Or ws01 > "12" Then
                MsgBox "返済年月日Toを入力してください"
                実行日2 = "": Call CEkey.SetFs(実行日2, True)
                Exit Sub
            End If
        
            実行日2 = 実行日2 & "01"
            
        ElseIf Len(実行日2) > 4 Then
            ws01 = Right$("000000" & 実行日2, 2)
            If ws01 < "01" Or ws01 > "31" Then
                MsgBox "返済年月日Toを入力してください"
                実行日2 = "": Call CEkey.SetFs(実行日2, True)
                Exit Sub
            End If
        End If
        
        実行日2 = C年月日.FormatDate("年月日", 実行日2)
        If C年月日.平成To西暦("年月日", 実行日2) = 0 Then
            MsgBox "返済年月日Toが違います"
            実行日2 = "": Call CEkey.SetFs(実行日2, True)
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
    Call MXA030_GRPTCLEAR
'
    wsSitei = P8.FCStr(銀行名.Text)
    wsShukei = P8.FCStr(集計区分.Text)
    
    ws01 = P8.FCStr(実行日)
    ws02 = P8.FCStr(実行日2)
    
    GVar1 = C年月日.平成To西暦("年月日", ws01)
    GVar2 = C年月日.平成To西暦("年月日", ws02)
    
    If ws01 <> "" And ws02 <> "" Then
        If CDate(GVar2) < CDate(GVar1) Then
            MsgBox "年月日が違います"
            実行日2 = "": Call CEkey.SetFs(実行日2, True)
            Exit Function
        End If
    ElseIf ws01 = "" And ws02 = "" Then
        MsgBox "年月日を入力してください"
        実行日 = "": Call CEkey.SetFs(実行日, True)
        Exit Function
        
    ElseIf ws01 = "" And ws02 <> "" Then
        MsgBox "年月日を入力してください"
        実行日 = "": Call CEkey.SetFs(実行日, True)
        Exit Function
    
    ElseIf ws01 <> "" And ws02 = "" Then
        MsgBox "年月日を入力してください"
        実行日2 = "": Call CEkey.SetFs(実行日2, True)
        Exit Function
    
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
    GRpt.テキスト_02 = 実行日2
    GRpt.集計 = wsShukei
    GRpt.指定 = wsSitei
    GRpt.詳細表示 = Check1
    GRpt.CSV = Check2
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

    Dim w実行日 As Date, wc実行日 As Date
'
    On Error GoTo 実行_Click_ERR
'
    ' =============================================================
    '       返済年月日Toの休日の場合翌営業にセットするか？
    ' =============================================================
    If IsDate(実行日2) Then
        w実行日 = C年月日.平成To西暦("年月日", 実行日2)
        Call C休日.計算(w実行日, 0)
        wc実行日 = C休日.次回稼働日
        
        If wc実行日 <> w実行日 Then
            GRet = MsgBox("返済年月日To(" & 実行日2 & ")は休日です。翌営業日にセットしますか？", vbYesNo + vbQuestion)
            If GRet = vbYes Then
                w実行日 = wc実行日
                実行日2 = Format(w実行日, Gfmt年月日)
            End If
        End If
    End If
'
    ' =========================================
    '           基本情報ファイル Read
    ' =========================================
    MAA010_基本情報ファイル_Read
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
    Set rpt = New RDA040_借入金返済予定表
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
    GLogStr = GLogStr & "返済年月日From=" & GRpt.テキスト_01 & ","
    GLogStr = GLogStr & "返済年月日To=" & GRpt.テキスト_02 & ","
    GLogStr = GLogStr & "銀行=" & GRpt.指定 & ","
    GLogStr = GLogStr & "集計=" & GRpt.集計 & ","
    GLogStr = GLogStr & "個別借入金表示=" & GRpt.詳細表示 & ","
    GLogStr = GLogStr & "千円単位=" & GRpt.千円単位 & ","
    GLogStr = GLogStr & "CSV出力=" & GRpt.CSV
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




