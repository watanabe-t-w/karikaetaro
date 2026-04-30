VERSION 5.00
Begin VB.Form frm_R利息前払未払明細表 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "利息明細表　出力"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10185
   Icon            =   "frm_R利息前払未払明細表.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox 借入番号 
      Height          =   300
      Left            =   2760
      TabIndex        =   0
      Text            =   "借入番号"
      Top             =   1080
      Width           =   3135
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
      Left            =   6240
      TabIndex        =   1
      Top             =   2280
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
      Left            =   8160
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CheckBox CSV出力 
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton 検索 
      Caption         =   "..."
      Height          =   300
      Left            =   5880
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
   End
   Begin VB.CheckBox 金利シミュレーション 
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2400
      Width           =   495
   End
   Begin VB.CheckBox 解約シミュレーション 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   2040
      Width           =   495
   End
   Begin 借換たろう.ZU050_Button ZU050_Button1 
      Height          =   615
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "利息明細表　出力"
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
      TabIndex        =   11
      Top             =   1080
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
      TabIndex        =   10
      Top             =   1440
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
      TabIndex        =   9
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   1  '右揃え
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  '実線
      Caption         =   "解約ｼﾐｭﾚｰｼｮﾝ"
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
      Top             =   2040
      Width           =   2535
   End
End
Attribute VB_Name = "frm_R利息前払未払明細表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "frm_R利息前払未払明細表"
'
Dim wRs As ADODB.Recordset
Dim wstr As String

'------------------------------------------------
' Form_Load
'------------------------------------------------
  Private Sub Form_Load()
'
    GRpt.帳票名 = "利息明細表"
    
    Me.Left = G_LEFT
    Me.Top = G_TOP
    
    借入番号 = ""
    '千円単位 = 0
    CSV出力 = 0
    金利シミュレーション = 0
    
    Call MFA001_借入番号(借入番号)
'    Call MFA001_金融リストラ番号(金融リストラ番号)

    'データ検索後
    If GStr_1 <> "" Then
'        借入番号 = GStr_1
    End If
'
    L_金利SM.Visible = False
    金利シミュレーション.Visible = False
    
    '金利GR
    If GStr = "金利GR" Then
        L_金利SM.Visible = True
        金利シミュレーション = 1
        金利シミュレーション.Visible = True
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

'------------------------------------------------
' 入力チェック
'------------------------------------------------
Private Function 入力チェック() As Boolean
'
    Dim FLG_Check As Boolean
    
    Dim ws01 As String, wsKin As String
'
    On Error GoTo 入力チェック_ERR
'
    入力チェック = False
    FLG_Check = False
    
    Call MXA030_GRPTCLEAR
'
    ws01 = P8.FCStr(借入番号)
    
    If 解約シミュレーション.Value = 1 Then
        '金融リストラ番号を取得
        wstr = ""
        wstr = wstr & "SELECT 金融リストラ番号"
        wstr = wstr & " FROM DBDA010_借入金"
        wstr = wstr & " WHERE 借入番号 = '" & ws01 & "'"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
            If Not wRs.EOF Then
                wsKin = P8.FCStr(wRs("金融リストラ番号"))
            Else
                wsKin = ""
            End If
        
        wRs.Close
        Set wRs = Nothing
    End If
    
    If ws01 = "" Then
        MsgBox "借入番号が未入力です", vbExclamation
        FLG_Check = True:
        Call CEkey.SetFs(借入番号, True)
    End If
        
    If FLG_Check = True Then
        Exit Function
    End If
'
    FLG_Check = False
'
    ' =========================================
    '                    後処理
    ' =========================================
    GRpt.コンボ_01 = 借入番号
    GRpt.金融 = wsKin
    
    'GRpt.千円単位 = 千円単位
    GRpt.CSV = CSV出力
    If 金利シミュレーション = 1 Then
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
    Set rpt = New RDA010_利息前払未払明細表
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
    GLogStr = GLogStr & "解約ｼﾐｭﾚｰｼｮﾝ=" & 解約シミュレーション.Value & ","
    GLogStr = GLogStr & "金利ｼﾐｭﾚｰｼｮﾝ=" & 金利シミュレーション.Value & ","
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


