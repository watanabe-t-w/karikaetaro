VERSION 5.00
Begin VB.Form frm_I固定項目登録 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "固定項目登録"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8880
   Icon            =   "frm_I固定項目登録.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox CSV日付書式 
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
      Left            =   2160
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   7
      Top             =   4080
      Width           =   2295
   End
   Begin VB.ComboBox 日付表示 
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
      Left            =   2160
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   6
      Top             =   3600
      Width           =   2295
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
      Left            =   6840
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton 保存 
      Caption         =   "保存（F11)"
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
      Left            =   4920
      TabIndex        =   8
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   9720
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9720
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9720
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9720
      Width           =   2535
   End
   Begin VB.TextBox 決算月 
      Alignment       =   1  '右揃え
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
      IMEMode         =   3  'ｵﾌ固定
      Left            =   2160
      MaxLength       =   2
      TabIndex        =   0
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox 決算日 
      Alignment       =   1  '右揃え
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
      IMEMode         =   3  'ｵﾌ固定
      Left            =   2160
      MaxLength       =   2
      TabIndex        =   2
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox 登録名称 
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
      IMEMode         =   4  '全角ひらがな
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   9
      Top             =   4560
      Width           =   6495
   End
   Begin 借換たろう.ZU020_ComboBox 借入金管理区分 
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Top             =   2160
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      ForeColor       =   -2147483640
      ForeColor       =   -2147483640
      IMEMode         =   3
      TextWidth       =   615
      BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      P8_ListBoxMax   =   0
      P8_KeySort      =   0   'False
   End
   Begin 借換たろう.ZU050_Button ZU050_Button1 
      Height          =   615
      Left            =   120
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "固定項目登録"
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
   Begin 借換たろう.ZU020_ComboBox 決算サイクル 
      Height          =   315
      Left            =   2160
      TabIndex        =   5
      Top             =   3120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      ForeColor       =   -2147483640
      ForeColor       =   -2147483640
      IMEMode         =   3
      TextWidth       =   615
      BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty P8_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      P8_ListBoxMax   =   0
      P8_KeySort      =   0   'False
   End
   Begin VB.Label Label11 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   "CSV日付書式"
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
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   " ※決算締日が末日の場合は、31と入力してください。"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   27
      Top             =   1680
      Width           =   5655
   End
   Begin VB.Label Label9 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   "決算サイクル"
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
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   "日付表示"
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
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   " ※決算用（支払日が銀行の休日の場合、翌営業にて算出）"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   16
      Top             =   2880
      Width           =   6975
   End
   Begin VB.Label Label3 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   " 下期決算月"
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
      Left            =   3120
      TabIndex        =   23
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   " 上期決算月"
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
      Left            =   3120
      TabIndex        =   22
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   " 決算月"
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
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label メッセージ 
      BackColor       =   &H00C0C000&
      Caption         =   "メッセージ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   9240
      Width           =   15015
   End
   Begin VB.Label L_上期決算月 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000000&
      BorderStyle     =   1  '実線
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
      Left            =   5160
      TabIndex        =   1
      Top             =   960
      Width           =   615
   End
   Begin VB.Label L_下期決算月 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000000&
      BorderStyle     =   1  '実線
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
      Left            =   5160
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label13 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   " 決算締日"
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
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label8 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   " 借入金管理区分"
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
      TabIndex        =   18
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   " ※管理用（予定の年月にて算出）"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   360
      TabIndex        =   17
      Top             =   2640
      Width           =   6975
   End
   Begin VB.Label Label6 
      Alignment       =   1  '右揃え
      BackColor       =   &H00D6DBBD&
      BorderStyle     =   1  '実線
      Caption         =   " 会社名"
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
      TabIndex        =   15
      Top             =   4560
      Width           =   2055
   End
End
Attribute VB_Name = "frm_I固定項目登録"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "固定項目登録"

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
    Me.Caption = GFcap
    Me.Left = G_LEFT
    Me.Top = G_TOP
'
    GStr = "": GStr_1 = "": GStr_2 = ""
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
'
    ' =========================================
    '             コンボボックス
    ' =========================================
    With 借入金管理区分
        .P8_Clear
        .P8_SqlString = ""
        .P8_KeyLeng = 2
        
        Call .AddItem(XMXA020_区分("借入金管理区分", "管理用"), "管理用")
        Call .AddItem(XMXA020_区分("借入金管理区分", "決算用"), "決算用")
    End With
    借入金管理区分.CreateCombo
    
    With 決算サイクル
        .P8_Clear
        .P8_SqlString = ""
        .P8_KeyLeng = 2
        
        Call .AddItem(XMXA020_区分("決算サイクル", "月次"), "月次")
        Call .AddItem(XMXA020_区分("決算サイクル", "四半期"), "四半期")
        Call .AddItem(XMXA020_区分("決算サイクル", "半期"), "半期")
        Call .AddItem(XMXA020_区分("決算サイクル", "年次"), "年次")
    End With
    決算サイクル.CreateCombo
    
    With 日付表示
        .Clear
        .AddItem "和暦", 0
        .AddItem "西暦", 1
        
        '旧ver 和暦チェック
        .Enabled = False
    End With
    
    With CSV日付書式
        .Clear
        .AddItem "YYYYMMDD", 0
        .AddItem "YYYY/MM/DD", 1
        .AddItem "YYYY-MM-DD", 2
        .AddItem "YYYY.MM.DD", 3
        .AddItem "YYMMDD", 4
        .AddItem "YY/MM/DD", 5
        .AddItem "YY-MM-DD", 6
        .AddItem "YY.MM.DD", 7
            '旧ver 和暦チェック
'        .AddItem "EEMMDD(和暦)", 8
'        .AddItem "EE/MM/DD(和暦)", 9
'        .AddItem "EE-MM-DD(和暦)", 10
'        .AddItem "EE.MM.DD(和暦)", 11
    End With
'
    Call 画面セット(False)
    メッセージ = ""
'
End Sub

'------------------------------------------------
' Form_KeyDown
'------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
    If KeyCode = vbKeyF11 Then
        Call 保存_Click
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
    メッセージ = ""
'
End Sub

'------------------------------------------------
' 画面セット
'------------------------------------------------
Private Function 画面セット(pGridClick As Boolean) As Boolean
'
    Dim j As Integer
    Dim ws01 As String
'
    On Error GoTo 画面セット_ERR
'
    画面セット = False
'
    ' =========================================
    '                画面クリア
    ' =========================================
    決算月 = ""
    決算日 = ""
    L_上期決算月.Caption = ""
    L_下期決算月.Caption = ""

    借入金管理区分.Text = ""
    決算サイクル.Text = ""
    
    登録名称 = ""
    
    ' =========================================
    '            マスタ セット
    ' =========================================
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAA010_基本情報"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.eof Then
            If 決算月 <> "" Then
                GRet = MsgBox("新規レコードを追加します。よろしいですか？", vbYesNo)
                If GRet = vbNo Then
                    wRs.Close
                    Set wRs = Nothing
                    
                    Exit Function
                End If
                Call CEkey.SetFs(決算月, True)
            End If
        Else
            画面セット = True
            
            Call CEkey.SetFs(決算月, True)
            決算月 = P8.FCStr(wRs("決算月"))
            決算日 = P8.FCStr(wRs("決算締日"))
            L_上期決算月.Caption = P8.FCStr(wRs("上期"))
            L_下期決算月.Caption = P8.FCStr(wRs("下期"))
            
            借入金管理区分.Text = P8.FCStr(wRs("借入金管理区分"))
            決算サイクル.Text = P8.FCStr(wRs("決算サイクル"))
            
            '旧ver 和暦チェック
            '日付表示.ListIndex = P8.FCDbl(wRs("日付入力区分"))
            'CSV日付書式.ListIndex = P8.FCDbl(wRs("CSV日付書式"))
            If P8.FCDbl(wRs("日付入力区分")) <> 1 Then
                日付表示.ListIndex = 0
            Else
                日付表示.ListIndex = P8.FCDbl(wRs("日付入力区分"))
            End If
            
            If P8.FCDbl(wRs("CSV日付書式")) > 7 Then
                CSV日付書式.ListIndex = 0
            Else
                CSV日付書式.ListIndex = P8.FCDbl(wRs("CSV日付書式"))
            End If

        
        End If
    wRs.Close
    Set wRs = Nothing
'
    wstr = ""
    wstr = wstr + "Select * From DAAA070_企業名マスタ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
        
        登録名称 = P8.FCStr(wRs("企業名"))
        
    End If
    wRs.Close
    Set wRs = Nothing
'
    If 登録名称 = "---------------" Then
        登録名称 = ""
    End If
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
' LostFocus
'------------------------------------------------
Private Sub 決算月_LostFocus()
'
    Dim wdate As Date, wDate2 As Date
    Dim ws01 As String
'
    If Not IsNumeric(決算月) Or (P8.FCDbl(決算月) > 12 Or P8.FCDbl(決算月) < 1) Then
        Exit Sub
    End If
    
    ws01 = "00" & P8.FCDbl(決算月)
    ws01 = Right$(ws01, 2)
    
    wdate = CDate("2001/" & ws01 & "/01")
    wDate2 = DateAdd("m", 6, wdate)
    
    L_上期決算月.Caption = Month(wDate2)
    L_下期決算月.Caption = P8.FCDbl(決算月)
'
End Sub

Private Sub 決算日_LostFocus()
    If Not IsNumeric(決算日) Or (P8.FCDbl(決算日) > 31 Or P8.FCDbl(決算日) < 1) Then
        Exit Sub
    End If
End Sub

Private Sub 登録名称_LostFocus()
    Call P8.FCControlLeft(登録名称, 30)
End Sub

'------------------------------------------------
' 保存_Click
'------------------------------------------------
Private Sub 保存_Click()
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

    '----------------------------------------
    '               登録チェック
    '----------------------------------------
    GRet = 登録_Check
    If GRet <> True Then
        Exit Sub
    End If
'
    Call 保存処理

    Call MAA010_基本LIST
'
    ' =========================================
    '               画面セット
    ' =========================================
    Call 画面セット(False)
    Call CEkey.SetFs(決算月, True)
'
End Sub

'------------------------------------------------
' 保存処理
'------------------------------------------------
Private Sub 保存処理()
'
    On Error GoTo 保存処理_ERR
'
    ' =========================================
    '            DAAA010_基本情報 更新処理
    ' =========================================
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DAAA010_基本情報"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.eof Then
            wRs.AddNew
        End If
     
            wRs("決算月") = P8.FCDbl(決算月)
            wRs("決算締日") = P8.FCDbl(決算日)
            wRs("上期") = P8.FCDbl(L_上期決算月.Caption)
            wRs("下期") = P8.FCDbl(L_下期決算月.Caption)
            
            wRs("借入金管理区分") = P8.FCDbl(借入金管理区分.Text)
            wRs("決算サイクル") = P8.FCDbl(決算サイクル.Text)
            
            wRs("日付入力区分") = P8.FCStr(日付表示.ListIndex)
            wRs("CSV日付書式") = P8.FCStr(CSV日付書式.ListIndex)
        
        wRs.Update
    wRs.Close
    Set wRs = Nothing
'
    GCoName = P8.FCStr(登録名称)
    'If GCoName = "" Then
    '    GCoName = "---------------"
    'End If
    
    wstr = ""
    wstr = wstr + "Select * From DAAA070_企業名マスタ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        wRs("企業名") = GCoName
        
        wRs.Update
    wRs.Close
    Set wRs = Nothing

    Call List_Set企業名

    If GCoName = "---------------" Then
        GCoName = ""
    End If
'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    GLogStr = "決算月=" & P8.FCDbl(決算月) & ","
    GLogStr = GLogStr & "決算締日=" & P8.FCDbl(決算日) & ","
    GLogStr = GLogStr & "上期=" & P8.FCDbl(L_上期決算月.Caption) & ","
    GLogStr = GLogStr & "下期=" & P8.FCDbl(L_下期決算月.Caption) & ","
    GLogStr = GLogStr & "借入金管理区分=" & P8.FCDbl(借入金管理区分.Text) & ","
    GLogStr = GLogStr & "決算サイクル=" & P8.FCDbl(決算サイクル.Text) & ","
    GLogStr = GLogStr & "日付表示=" & P8.FCStr(日付表示.Text) & ","
    GLogStr = GLogStr & "CSV日付書式=" & P8.FCStr(CSV日付書式.Text) & ","
    GLogStr = GLogStr & "企業名=" & GCoName
    Call MXA030_LOG_WRITE("固定項目登録", "更新", GLogStr)
'
    ' =========================================
    '                テーブル変更
    ' =========================================
    Call MAA010_基本情報ファイル_Read
'
    ' =========================================
    '         DAAA030_科目マスタ 設定
    ' =========================================
    Call MAA500_科目コード設定
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "保存しました。すべてのフォームを閉じます。", vbInformation
    Call UNLOAD_ALLFRM
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
保存処理_ERR:
    pERR_MES = pPROGRAM_ID + "/ 保存処理() でエラー" + vbCrLf + vbCrLf + _
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
' 登録_Check
'------------------------------------------------
Private Function 登録_Check() As Boolean
'
    Dim wi01 As Integer
'
    登録_Check = False
'
    On Error GoTo 登録_Check_ERR
'
    wi01 = P8.FCDbl(決算月)
    If Not IsNumeric(決算月) Or (wi01 > 12 Or wi01 < 1) Then
        MsgBox "入力を確認してください": Call CEkey.SetFs(決算月, True)
        Exit Function
    End If
'
    wi01 = P8.FCDbl(決算日)
    If Not IsNumeric(決算日) Or (wi01 > 31 Or wi01 < 1) Then
        MsgBox "入力を確認してください": Call CEkey.SetFs(決算日, True)
        Exit Function
    End If
'

'コンボ
    If Not IsNumeric(借入金管理区分.Text) Then
        MsgBox "入力を確認してください": Call CEkey.SetFs(借入金管理区分, True)
        Exit Function
    End If
    If P8.FCDbl(借入金管理区分.Text) <> "0" And P8.FCDbl(借入金管理区分.Text) <> "1" Then
        MsgBox "借入金管理区分を確認してください": Call CEkey.SetFs(借入金管理区分, True)
        Exit Function
    End If
'
    If Not IsNumeric(決算サイクル.Text) Then
        MsgBox "入力を確認してください": Call CEkey.SetFs(決算サイクル, True)
        Exit Function
    End If
    If P8.FCDbl(決算サイクル.Text) <> "1" And P8.FCDbl(決算サイクル.Text) <> "3" _
    And P8.FCDbl(決算サイクル.Text) <> "6" And P8.FCDbl(決算サイクル.Text) <> "12" Then
        MsgBox "決算サイクルを確認してください": Call CEkey.SetFs(決算サイクル, True)
        Exit Function
    End If
'
    '旧ver 和暦設定CHECK
    If P8.FCDbl(日付表示.ListIndex) <> 1 Then
        MsgBox "西暦を選択してください": Call CEkey.SetFs(日付表示, True)
        Exit Function
    End If
    
    If P8.FCDbl(CSV日付書式.ListIndex) > 7 Then
        MsgBox "CSV日付書式を確認してください": Call CEkey.SetFs(CSV日付書式, True)
        Exit Function
    End If
'
    登録_Check = True
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
登録_Check_ERR:
    pERR_MES = pPROGRAM_ID + "/ 登録_Check() でエラー" + vbCrLf + vbCrLf + _
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
' List_Set企業名
'------------------------------------------------
Private Sub List_Set企業名()
'
    Dim wDb As New ADODB.Connection
    Dim wRs2 As ADODB.Recordset
'
    On Error GoTo List_Set企業名_ERR
'
    '----------< List.mdb Open >------------------------------------------------
    Call AdoDbOpen("Jet", wDb, GCurDir + "\" + GMain, "", , GPwd)

    wstr = ""
    wstr = wstr + "Select * From DAAA070_企業名マスタ"
    wstr = wstr + " Where 企業名Key='" & GKeyName & "'"
    Call AdoRecordsetOpen(wDb, wRs2, wstr)
    If Not wRs2.eof Then
        wRs2("企業名") = GCoName
        
        wRs2.Update
    End If
    wRs2.Close
    Set wRs2 = Nothing
    
    wDb.Close
    Set wDb = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
List_Set企業名_ERR:
    pERR_MES = pPROGRAM_ID + "/ List_Set企業名() でエラー" + vbCrLf + vbCrLf + _
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
    Unload Me
End Sub


