VERSION 5.00
Begin VB.Form frm_R利息未払前払残高表 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "利息未払前払残高表　出力"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10485
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox 実行日 
      Alignment       =   1  '右揃え
      Height          =   330
      IMEMode         =   2  'ｵﾌ
      Left            =   2760
      TabIndex        =   1
      Text            =   "HH年MM月"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox 実行日2 
      Alignment       =   1  '右揃え
      Height          =   330
      IMEMode         =   2  'ｵﾌ
      Left            =   2760
      TabIndex        =   2
      Text            =   "HH年"
      Top             =   1800
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
      Left            =   6600
      TabIndex        =   5
      Top             =   4680
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
      Left            =   8520
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1815
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
   Begin VB.ComboBox 金融リストラ番号 
      Height          =   300
      IMEMode         =   1  'ｵﾝ
      Left            =   2760
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   3
      Top             =   2280
      Width           =   3135
   End
   Begin VB.ComboBox 銀行名 
      Height          =   300
      IMEMode         =   1  'ｵﾝ
      Left            =   2760
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   4
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CheckBox Check2 
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   3480
      Width           =   495
   End
   Begin VB.CheckBox Check4 
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   4320
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   3120
      Width           =   495
   End
   Begin VB.CheckBox Check3 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   3840
      Width           =   495
   End
   Begin 借換たろう.ZU050_Button ZU050_Button1 
      Height          =   615
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   5775
      _ExtentX        =   8281
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "利息未払前払残高表　出力"
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
   Begin VB.Label L_実行日 
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
      TabIndex        =   20
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label L_実行日2 
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
      TabIndex        =   19
      Top             =   1800
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
      TabIndex        =   18
      Top             =   960
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
      TabIndex        =   17
      Top             =   2280
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
      TabIndex        =   16
      Top             =   2640
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
      TabIndex        =   15
      Top             =   3480
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
      TabIndex        =   14
      Top             =   4320
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
      TabIndex        =   13
      Top             =   3120
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
      TabIndex        =   12
      Top             =   3840
      Width           =   2535
   End
End
Attribute VB_Name = "frm_R利息未払前払残高表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "frm_R利息未払前払残高表"
'
Dim wRs As ADODB.Recordset
Dim wstr As String

'------------------------------------------------
' Form_Load
'------------------------------------------------
  Private Sub Form_Load()
'
    GRpt.帳票名 = "利息未払前払残高表"
    
    Me.Left = G_LEFT
    Me.Top = G_TOP
    
    実行日 = ""
    実行日2 = ""
    出力表種別.Clear
    金融リストラ番号.Clear
    銀行名.Clear
    Check1 = 0
    Check2 = 0
    Check3 = 0
    Check4 = 0
    
    Call MFA001_出力表種別(出力表種別)
    Call MFA001_金融リストラ番号(金融リストラ番号)
    Call MFA001_銀行区分(銀行名)
'
    L_金利SM.Visible = False
    Check4.Visible = False
    
    '金利GR
    If GStr = "金利GR" Then
        L_金利SM.Visible = True
        Check4 = 1
        Check4.Visible = True
    End If
'
    実行日 = Format(Now, Gfmt年月)
    実行日2 = Replace(Format(C年月日.年度開始年月日(C年月日.年度変換(CStr(Now)), "西暦"), Gfmt年), "年", "")

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
    Dim wsSuii As String
    Dim j As Integer
    Dim wJikou As Date, wdate As Date
    Dim FLG_DATE As Boolean
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
'
    wsSuii = P8.FCStr(出力表種別.Text)
    FLG_DATE = False
    If wsSuii = "四半期" Then
        wJikou = C年月日.平成To西暦("年月", 実行日)
        wdate = C年月日.年度開始年月日(C年月日.年度変換(CStr(wJikou)), "西暦")
        wdate = DateAdd("m", -1, wdate)
        For j = 1 To 12
            If wJikou = Format(wdate, "yyyy/mm") Then
                FLG_DATE = True
                    Exit For
            End If
            wdate = DateAdd("m", 3, wdate)
        Next
        
        If FLG_DATE = False Then
            MsgBox "年月が違います"
            実行日 = "": Call CEkey.SetFs(実行日, True)
            Exit Sub
        End If
    ElseIf wsSuii = "半期" Then
        wJikou = C年月日.平成To西暦("年月", 実行日)
        wdate = C年月日.年度開始年月日(C年月日.年度変換(CStr(wJikou)), "西暦")
        wdate = DateAdd("m", -1, wdate)
        For j = 1 To 12
            If wJikou = Format(wdate, "yyyy/mm") Then
                FLG_DATE = True
                    Exit For
            End If
            wdate = DateAdd("m", 6, wdate)
        Next
        
        If FLG_DATE = False Then
            MsgBox "年月が違います"
            実行日 = "": Call CEkey.SetFs(実行日, True)
            Exit Sub
        End If
    End If
'
    Call CEkey.AllSelect
'
End Sub

Private Sub 出力表種別_Click()
'
    Dim wsSuii As String
    Dim j As Integer
    Dim wJikou As Date, wdate As Date
'
    L_実行日.Visible = True
    実行日.Visible = True
    L_実行日2.Visible = False
    実行日2.Visible = False
    If 出力表種別 = "年次" Then
        L_実行日.Visible = False
        実行日.Visible = False
        L_実行日2.Visible = True
        実行日2.Visible = True
    Else
        L_実行日.Visible = True
        実行日.Visible = True
        L_実行日2.Visible = False
        実行日2.Visible = False
    
        wsSuii = P8.FCStr(出力表種別.Text)
        If wsSuii = "四半期" Then
            wJikou = C年月日.平成To西暦("年月", Format(Now, Gfmt年月))
            wdate = C年月日.年度開始年月日(C年月日.年度変換(CStr(wJikou)), "西暦")
            wdate = DateAdd("m", -1, wdate)
            wdate = DateAdd("yyyy", 1, wdate)
            For j = 12 To 1 Step -1
                If wJikou >= Format(wdate, "yyyy/mm") Then
                    実行日 = Format(wdate, Gfmt年月)
                        Exit For
                End If
                wdate = DateAdd("m", -3, wdate)
            Next
        ElseIf wsSuii = "半期" Then
            wJikou = C年月日.平成To西暦("年月", Format(Now, Gfmt年月))
            wdate = C年月日.年度開始年月日(C年月日.年度変換(CStr(wJikou)), "西暦")
            wdate = DateAdd("m", -1, wdate)
            wdate = DateAdd("yyyy", 1, wdate)
            For j = 12 To 1 Step -1
                If wJikou >= Format(wdate, "yyyy/mm") Then
                    実行日 = Format(wdate, Gfmt年月)
                        Exit For
                End If
                wdate = DateAdd("m", -6, wdate)
            Next
        End If
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
    Dim wsUri As String, wsUri2 As String
    Dim wsKar As String, wsKar2 As String
    Dim wsSet As String, wsSet2 As String
    Dim wsKin As String, wskin2 As String
    Dim wsSeR As String, wsSeR2 As String
    Dim wsLea As String, wsLea2 As String
'
    On Error GoTo 入力チェック_ERR
'
    '
    GRpt.推移 = ""
    GRpt.選択 = ""
    GRpt.実績 = ""
    GRpt.作業 = ""
    GRpt.集計 = ""
    GRpt.指定 = ""
    
    GRpt.連結売上 = ""
    GRpt.売上 = ""
    GRpt.借入 = ""
    GRpt.設備 = ""
    GRpt.金融 = ""
    GRpt.設備R = ""
    GRpt.リス = ""
    
    GRpt.連結売上2 = ""
    GRpt.売上2 = ""
    GRpt.借入2 = ""
    GRpt.設備2 = ""
    GRpt.金融2 = ""
    GRpt.設備R2 = ""
    GRpt.リス2 = ""
    
    GRpt.テキスト_01 = ""
    GRpt.テキスト_02 = ""
    
    GRpt.借入金管理区分 = 0
    GRpt.詳細表示 = 0
    GRpt.CSV = 0
    GRpt.千円単位 = 0
    GRpt.金利SM = 0
      
    GRpt.チェック_01 = 0
    GRpt.チェック_02 = 0
    GRpt.チェック_03 = 0
    GRpt.チェック_04 = 0

    G金利SM = False
'
    wsSuii = P8.FCStr(出力表種別.Text)
    wsKin = P8.FCStr(金融リストラ番号.Text)
    wsSitei = P8.FCStr(銀行名.Text)
    
    If wsSuii <> "年次" Then
        ws01 = P8.FCStr(実行日)
        GVar1 = C年月日.平成To西暦("年月", ws01)
    
        If ws01 = "" Then
            MsgBox "年月を入力してください"
            実行日 = "": Call CEkey.SetFs(実行日, True)
            Exit Function
        
        End If
    
    Else
        ws01 = P8.FCStr(実行日2)
        If Not P8.FIsInt(実行日2) Then
            実行日2 = ""
            FLG_Check = True:   Call CEkey.SetFs(実行日2, True)
        End If
        
        If P8.FCDbl(実行日2) <= 0 Then
            実行日 = ""
            FLG_Check = True:   Call CEkey.SetFs(実行日2, True)
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
    GRpt.テキスト_02 = 実行日2
    GRpt.推移 = wsSuii
    GRpt.集計 = wsShukei
    GRpt.指定 = wsSitei
    GRpt.金融 = wsKin
    GRpt.詳細表示 = Check1
    GRpt.千円単位 = Check2
    GRpt.CSV = Check3
    If Check4 = 1 Then
        G金利SM = True
    End If
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
    Set rpt = New RDF010_利息未払前払残高表
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
    GLogStr = "出力表種別=" & GRpt.推移 & ","
    GLogStr = GLogStr & "年月=" & GRpt.テキスト_01 & ","
    GLogStr = GLogStr & "年度=" & GRpt.テキスト_02 & ","
    GLogStr = GLogStr & "ｼﾐｭﾚｰｼｮﾝ番号=" & GRpt.金融 & ","
    GLogStr = GLogStr & "集計区分=" & GRpt.集計 & ","
    GLogStr = GLogStr & "銀行名=" & GRpt.指定 & ","
    GLogStr = GLogStr & "個別借入金表示=" & GRpt.詳細表示 & ","
    GLogStr = GLogStr & "千円単位=" & GRpt.千円単位 & ","
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

