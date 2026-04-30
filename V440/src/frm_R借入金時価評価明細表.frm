VERSION 5.00
Begin VB.Form frm_R借入金時価評価明細表 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "借入金時価評価明細表　出力"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10035
   Icon            =   "frm_R借入金時価評価明細表.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox 実行日 
      Alignment       =   1  '右揃え
      Height          =   330
      IMEMode         =   2  'ｵﾌ
      Left            =   1560
      TabIndex        =   1
      Text            =   "HH年"
      Top             =   1440
      Width           =   735
   End
   Begin VB.ComboBox 実行月 
      Height          =   300
      Left            =   1560
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton 検索 
      Caption         =   "..."
      Height          =   300
      Left            =   4680
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
   End
   Begin VB.CheckBox CSV出力 
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   3000
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3120
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
      TabIndex        =   6
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CheckBox 千円単位 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   2640
      Width           =   495
   End
   Begin VB.ComboBox 借入番号 
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Text            =   "借入番号"
      Top             =   1080
      Width           =   3135
   End
   Begin 借換たろう.ZU050_Button ZU050_Button1 
      Height          =   615
      Left            =   240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   5775
      _ExtentX        =   8916
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "借入金時価評価明細表　出力"
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
   Begin VB.Label L_決算日 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "HH年MM月DD日"
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
      Left            =   1560
      TabIndex        =   3
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "年"
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
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label5 
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
      TabIndex        =   14
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "決算日"
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
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  '実線
      Caption         =   "借入番号"
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
      Top             =   1080
      Width           =   1335
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
      TabIndex        =   11
      Top             =   3000
      Width           =   2535
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
      TabIndex        =   10
      Top             =   2640
      Width           =   2535
   End
End
Attribute VB_Name = "frm_R借入金時価評価明細表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "frm_R借入金時価評価明細表"
'
Dim wRs As ADODB.Recordset
Dim wstr As String

'------------------------------------------------
' Form_Load
'------------------------------------------------
  Private Sub Form_Load()
'
    GRpt.帳票名 = "借入金時価評価明細表"
    
    Me.Left = G_LEFT
    Me.Top = G_TOP
    
    借入番号 = ""
    千円単位 = 0
    CSV出力 = 0
    実行日 = ""

    
    Call MFA001_借入番号_時価評価(借入番号)

    'データ検索後
    If GStr_1 <> "" Then
'        借入番号 = GStr_1
    End If
'
    '実行日 = Replace(Format(C年月日.年度開始年月日(C年月日.年度変換(CStr(Now)), "西暦"), Gfmt年), "年", "")
    実行日 = Replace(Format(Now, Gfmt年), "年", "")
'
    Call 実行月_セット
'
    Call 決算日_セット 'ADD 20170501 M.Mino
    
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
' 借入番号_LostFocus
'------------------------------------------------
Private Sub 借入番号_LostFocus()
    Call 決算日_セット
End Sub

'------------------------------------------------
' 実行日_LostFocus
'------------------------------------------------
Private Sub 実行日_LostFocus()

    'ADD START 20170501 M.Mino
    If IsNumeric(実行日) = False Then
        実行日 = ""
        Exit Sub
    End If
    'ADD END 20170501

    If G基本情報.日付入力区分 = "0" Then
    '和暦
        Call P8.FCControlLeft(実行日, 2)
    Else
    '西暦
        Call P8.FCControlLeft(実行日, 4)
    End If
'
    Call CEkey.AllSelect
    Call 決算日_セット
'
End Sub

'------------------------------------------------
' 実行月_LostFocus
'------------------------------------------------
Private Sub 実行月_LostFocus()

    'ADD START 20170501 M.Mino
    If IsNumeric(実行日) = False Then
        実行日 = ""
        Exit Sub
    End If
    If IsNumeric(実行月) = False Then
        実行日 = ""
        Exit Sub
    End If
    'ADD END 20170501

    Call 決算日_セット
    
End Sub

'------------------------------------------------
' 決算日_LostFocus
'------------------------------------------------
Private Sub 決算日_LostFocus()

    L_決算日.Caption = C年月日.FormatDate("年月日", L_決算日.Caption)

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
' 決算日_セット
'------------------------------------------------
Private Sub 決算日_セット()
'
    Dim ws01 As String, ws02 As String

    'ADD START 20170501 M.Mino
    If IsNumeric(実行日) = False Then
        Exit Sub
    End If
    If IsNumeric(実行月) = False Then
        Exit Sub
    End If
    'ADD END 20170501

    ws01 = "": ws02 = ""
    If G基本情報.日付入力区分 = "0" Then
    '和暦
        ws01 = 実行日 & "年" & Right(CStr("00" & 実行月), 2) & "月"
    Else
    '西暦
        ws01 = 実行日 & "/" & Right(CStr("00" & 実行月), 2)
    End If
    
    ws02 = C年月日.平成To西暦("年月日", ws01)
    ws02 = MBA010_締日年月日(P8.FCDate(ws02))

    L_決算日.Caption = Format(ws02, Gfmt年月日)
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

'------------------------------------------------
' 入力チェック
'------------------------------------------------
Private Function 入力チェック() As Boolean
'
    Dim j As Integer
    Dim FLG_Check As Boolean
    
    Dim ws01 As String, ws02 As String
'
    On Error GoTo 入力チェック_ERR
'
    Call MXA030_GRPTCLEAR
'
    If 借入番号.Text = "" Then
        MsgBox "借入番号を入力してください"
        Call CEkey.SetFs(借入番号, True): Exit Function
    End If
    
    If Not P8.FIsInt(実行日) Then
        実行日 = ""
        FLG_Check = True:   Call CEkey.SetFs(実行日, True)
        'ADD 20170501 M.Mino
        MsgBox ("年を入力してください")
        Exit Function
    End If
    
    If P8.FCDbl(実行日) <= 0 Then
        実行日 = ""
        FLG_Check = True:   Call CEkey.SetFs(実行日, True)
        'ADD 20170501 M.Mino
        MsgBox ("月を入力してください")
        Exit Function
    End If
    
    ws01 = ""
    If G基本情報.日付入力区分 = "0" Then
    '和暦
        ws01 = 実行日 & "年" & Right(CStr("00" & 実行月), 2) & "月"
    Else
    '西暦
        ws01 = 実行日 & "/" & Right(CStr("00" & 実行月), 2)
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
        
    GVar1 = C年月日.平成To西暦("年月日", L_決算日.Caption)
    If GVar1 = 0 Then
        FLG_Check = True: Call CEkey.SetFs(実行月, True)
    Else
        G決算日(1) = P8.FCDate(GVar1)
    End If
    
    If FLG_Check = True Then
        MsgBox "指定された内容に誤りがあります"

        Exit Function
    End If
'
    FLG_Check = False

    ' =========================================
    '                    後処理
    ' =========================================
    GRpt.コンボ_01 = 借入番号
    GRpt.テキスト_01 = G決算日(1)

    GRpt.千円単位 = 千円単位
    GRpt.CSV = CSV出力
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
' 検索_Click
'------------------------------------------------
Private Sub 検索_Click()
'
    GStr = GRpt.帳票名
    GStr_1 = ""
'
    Me.Enabled = False
'
    frm_K借入金検索.Show
'
End Sub


'------------------------------------------------
' 実行_Click
'------------------------------------------------
Private Sub 実行_Click()
'
    Dim rpt As Object
    
    Dim wIndex As Integer, j As Integer
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
    GRpt.帳票名 = "借入金時価評価明細表"
    Set rpt = New REA010_借入金時価評価明細表
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

'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    GLogStr = "借入番号=" & GRpt.コンボ_01 & ","
    GLogStr = GLogStr & "千円単位=" & GRpt.千円単位 & ","
    GLogStr = GLogStr & "CSV出力=" & GRpt.CSV & ","
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
' 閉じる_Click
'------------------------------------------------
Private Sub 閉じる_Click()
'
    Call MXA030_MCLEAR
'
    Unload Me
'
End Sub
