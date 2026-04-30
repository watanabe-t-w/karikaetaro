VERSION 5.00
Begin VB.Form frm_R帳票出力設定 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "帳票出力設定"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   Icon            =   "frm_R帳票出力設定.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton 実行 
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
      Left            =   4320
      TabIndex        =   51
      Top             =   5640
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
      Left            =   6240
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Frame Frame_Risoku 
      Height          =   675
      Left            =   240
      TabIndex        =   40
      Top             =   4680
      Width           =   7815
      Begin VB.OptionButton OpR 
         Caption         =   "表示しない"
         Height          =   255
         Index           =   4
         Left            =   5040
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox CheckR 
         Caption         =   "改ページ"
         Height          =   255
         Left            =   6480
         TabIndex        =   47
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OpR 
         Caption         =   "分類1"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpR 
         Caption         =   "分類2"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpR 
         Caption         =   "分類3"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpR 
         Caption         =   "分類4"
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.Label L_利息区分 
         Caption         =   "利息区分"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   7815
      Begin VB.CheckBox CheckS 
         Caption         =   "改ページ"
         Height          =   255
         Left            =   6480
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OpS 
         Caption         =   "分類1"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpS 
         Caption         =   "分類2"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpS 
         Caption         =   "分類3"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpS 
         Caption         =   "分類4"
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpS 
         Caption         =   "表示しない"
         Height          =   255
         Index           =   4
         Left            =   5040
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label L_借入金種別 
         Caption         =   "借入金種別"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   600
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   7815
      Begin VB.CheckBox CheckB 
         Caption         =   "改ページ"
         Height          =   255
         Left            =   6480
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OpB 
         Caption         =   "分類1"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpB 
         Caption         =   "分類2"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpB 
         Caption         =   "分類3"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpB 
         Caption         =   "分類4"
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpB 
         Caption         =   "表示しない"
         Height          =   255
         Index           =   4
         Left            =   5040
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label L_部門 
         Caption         =   "部門"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Height          =   600
      Left            =   240
      TabIndex        =   24
      Top             =   3480
      Width           =   7815
      Begin VB.CheckBox CheckG 
         Caption         =   "改ページ"
         Height          =   255
         Left            =   6480
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OpG 
         Caption         =   "分類1"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpG 
         Caption         =   "分類2"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpG 
         Caption         =   "分類3"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpG 
         Caption         =   "分類4"
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpG 
         Caption         =   "表示しない"
         Height          =   255
         Index           =   4
         Left            =   5040
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label L_銀行番号 
         Caption         =   "銀行番号"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Height          =   600
      Left            =   240
      TabIndex        =   16
      Top             =   2880
      Width           =   7815
      Begin VB.CheckBox CheckK 
         Caption         =   "改ページ"
         Height          =   255
         Left            =   6480
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OpK 
         Caption         =   "分類1"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpK 
         Caption         =   "分類2"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpK 
         Caption         =   "分類3"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpK 
         Caption         =   "分類4"
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpK 
         Caption         =   "表示しない"
         Height          =   255
         Index           =   4
         Left            =   5040
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label L_金融機関 
         Caption         =   "金融機関"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame_KSM 
      Height          =   600
      Left            =   240
      TabIndex        =   32
      Top             =   4080
      Width           =   7815
      Begin VB.CheckBox CheckM 
         Caption         =   "改ページ"
         Height          =   255
         Left            =   6480
         TabIndex        =   39
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OpM 
         Caption         =   "分類1"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpM 
         Caption         =   "分類2"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpM 
         Caption         =   "分類3"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpM 
         Caption         =   "分類4"
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   240
         Width           =   860
      End
      Begin VB.OptionButton OpM 
         Caption         =   "表示しない"
         Height          =   255
         Index           =   4
         Left            =   5040
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label L_金利シミュレーションG 
         Caption         =   "金利ｼﾐｭﾚｰｼｮﾝG"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1455
      End
   End
   Begin 借換たろう.ZU050_Button ZU050_Button1 
      Height          =   615
      Left            =   120
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "帳票出力設定"
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
   Begin VB.Image Image1 
      Height          =   450
      Left            =   7200
      Picture         =   "frm_R帳票出力設定.frx":0ECA
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H00C0FFFF&
      Caption         =   " 帳票名"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   240
      TabIndex        =   53
      Top             =   960
      Width           =   855
   End
   Begin VB.Label L_帳票名 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1080
      TabIndex        =   52
      Top             =   960
      Width           =   6975
   End
   Begin VB.Label L_集計分類 
      Caption         =   "集計分類（大 → 中 → 小）"
      Height          =   315
      Left            =   240
      TabIndex        =   48
      Top             =   1440
      Width           =   4575
   End
End
Attribute VB_Name = "frm_R帳票出力設定"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "frm_R帳票出力設定"
'
Dim wRs As ADODB.Recordset
Dim wstr As String

'------------------------------------------------
' Form_Load
'------------------------------------------------
Private Sub Form_Load()
'
    Dim j As Integer
'
    Me.Caption = GFcap
    Me.Left = G_LEFT
    Me.Top = G_TOP
'
    '分類集計
    OpS(0).Value = False: OpS(1).Value = False: OpS(2).Value = False: OpS(3).Value = False: OpS(4).Value = True
    OpB(0).Value = False: OpB(1).Value = False: OpB(2).Value = False: OpB(3).Value = False: OpB(4).Value = True
    OpK(0).Value = False: OpK(1).Value = False: OpK(2).Value = False: OpK(3).Value = False: OpK(4).Value = True
    OpG(0).Value = False: OpG(1).Value = False: OpG(2).Value = False: OpG(3).Value = False: OpG(4).Value = True
    OpM(0).Value = False: OpM(1).Value = False: OpM(2).Value = False: OpM(3).Value = False: OpM(4).Value = True
    
    For j = 0 To 4
        OpS(j).Visible = True
        OpB(j).Visible = True
        OpK(j).Visible = True
        OpG(j).Visible = True
        OpM(j).Visible = True
        
        OpR(j).Visible = False
    Next j
'
    Frame1.Visible = True
    Frame2.Visible = True
    Frame3.Visible = True
    Frame4.Visible = True
    Frame_Risoku.Visible = False
    Frame_KSM.Visible = False
'
    '改ページ
    CheckS = 1
    CheckB = 0
    CheckK = 0
    CheckG = 0
    CheckR = 0
    CheckM = 0

    CheckS.Visible = True
    CheckB.Visible = True
    CheckK.Visible = True
    CheckG.Visible = True
    CheckR.Visible = True
    CheckM.Visible = True
'

    OpR(0).Value = False: OpR(1).Value = False: OpR(2).Value = False: OpR(3).Value = False: OpR(4).Value = True
    If GForm.Name = "frm_R利息前払未払残高表" Then
        OpR(0).Value = True: OpR(1).Value = False: OpR(2).Value = False: OpR(3).Value = False: OpR(4).Value = False
        OpR(0).Visible = True
    End If
    If GForm.Name = "frm_R利息前払未払残高表" Then
        Frame_Risoku.Visible = True
    End If
'
    '杉村倉庫仕様
    If GForm.Name = "frm_R銀行別利息表" Then
        OpR(0).Value = True: OpR(1).Value = False: OpR(2).Value = False: OpR(3).Value = False: OpR(4).Value = False
        OpR(0).Visible = True
    End If
    If GForm.Name = "frm_R支払利息推移表" Then
        For j = 0 To 4
            OpR(j).Visible = True
        Next j
    End If
    
    L_部門.Enabled = True
    L_金融機関.Enabled = True
    If GForm.Name = "frm_R銀行別利息表" _
    Or GForm.Name = "frm_R支払利息推移表" Then
        Frame2.Enabled = False
        Frame3.Enabled = False
        Frame_Risoku.Visible = True
        L_部門.Enabled = False
        L_金融機関.Enabled = False
    End If
    If GForm.Name = "frm_R1年内返済集計表" Then
        Frame2.Enabled = False
        Frame3.Enabled = False
        L_部門.Enabled = False
        L_金融機関.Enabled = False
    End If
'
'
    L_帳票名.Caption = ""
'
    Call 画面セット
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    
    GForm.Enabled = True
    GForm.Show
End Sub

'------------------------------------------------
' 画面セット
'------------------------------------------------
Private Sub 画面セット()
'
    Dim wTable As MAA070_帳票出力設定
    Dim j As Integer
'
    wTable = MAA070_帳票出力設定Read(GForm.Name)

    L_帳票名.Caption = wTable.帳票名
'
    '分類1～4
    For j = 0 To 3
        If wTable.B_種別 = j + 1 Then
            OpS(j).Value = True
        End If
        If wTable.B_部門 = j + 1 Then
            OpB(j).Value = True
        End If
        If wTable.B_金融 = j + 1 Then
            OpK(j).Value = True
        End If
        If wTable.B_銀行 = j + 1 Then
            OpG(j).Value = True
        End If
        If wTable.B_金利 = j + 1 Then
            OpM(j).Value = True
        End If
        If wTable.B_利息 = j + 1 Then
            OpR(j).Value = True
        End If
    Next j
    
    '表示しない
    If wTable.B_種別 = 9 Then
        OpS(4).Value = True
    End If
    If wTable.B_部門 = 9 Then
        OpB(4).Value = True
    End If
    If wTable.B_金融 = 9 Then
        OpK(4).Value = True
    End If
    If wTable.B_銀行 = 9 Then
        OpG(4).Value = True
    End If
    If wTable.B_金利 = 9 Then
        OpM(4).Value = True
    End If
    If wTable.B_利息 = 9 Then
        OpR(4).Value = True
    End If

    If GForm.Name = "frm_R利息前払未払残高表" Then
        OpR(0).Value = True
    End If
    
    '杉村倉庫仕様
    If GForm.Name = "frm_R銀行別利息表" Then
        OpR(0).Value = True
    End If
    
    If GStr <> "金利GR" Then
        OpM(4).Value = True
    End If
'
    CheckS = wTable.P_種別
    CheckB = wTable.P_部門
    CheckK = wTable.P_金融
    CheckG = wTable.P_銀行
    CheckR = wTable.P_利息
    CheckM = wTable.P_金利
'
    If GStr = "金利GR" Then
        Frame_KSM.Visible = True
    End If
'
    If GForm.Name = "frm_R借入一覧表" Or GForm.Name = "frm_R年度別比較表" Then
        For j = 2 To 3
            OpS(j).Visible = False
            OpB(j).Visible = False
            OpK(j).Visible = False
            OpG(j).Visible = False
            OpM(j).Visible = False
            OpR(j).Visible = False
        Next j
    End If
    
    If GForm.Name = "frm_R金融機関別残高表" Then
        Frame1.Visible = False
        Frame2.Visible = False
        
        CheckK.Visible = False
        CheckG.Visible = False
        
        For j = 1 To 3
            OpS(j).Visible = False
            OpB(j).Visible = False
            OpK(j).Visible = False
            OpG(j).Visible = False
            OpM(j).Visible = False
            OpR(j).Visible = False
        Next j
    End If
'
End Sub

'------------------------------------------------
' OptionButton
'------------------------------------------------
Private Sub OpR_Click(Index As Integer)
    Call 分類OPOFF(Index, OpR(Index))
End Sub
Private Sub OpS_Click(Index As Integer)
    Call 分類OPOFF(Index, OpS(Index))
End Sub
Private Sub OpB_Click(Index As Integer)
    Call 分類OPOFF(Index, OpB(Index))
End Sub
Private Sub OpK_Click(Index As Integer)
    Call 分類OPOFF(Index, OpK(Index))
End Sub
Private Sub OpG_Click(Index As Integer)
    Call 分類OPOFF(Index, OpG(Index))
End Sub
Private Sub OpM_Click(Index As Integer)
    Call 分類OPOFF(Index, OpM(Index))
End Sub

'------------------------------------------------
' 分類OPOFF
'------------------------------------------------
Private Sub 分類OPOFF(pIndex As Integer, pOp As Control)
    If OpS(pIndex).Value = True And pOp.Name <> "OpS" Then
        OpS(4).Value = True
    ElseIf OpB(pIndex).Value = True And pOp.Name <> "OpB" Then
        OpB(4).Value = True
    ElseIf OpK(pIndex).Value = True And pOp.Name <> "OpK" Then
        OpK(4).Value = True
    ElseIf OpG(pIndex).Value = True And pOp.Name <> "OpG" Then
        OpG(4).Value = True
    ElseIf OpM(pIndex).Value = True And pOp.Name <> "OpM" Then
        OpM(4).Value = True
    ElseIf OpR(pIndex).Value = True And pOp.Name <> "OpR" Then
        If GForm.Name <> "frm_R利息前払未払残高表" Then
            OpR(4).Value = True
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
'
    On Error GoTo 入力チェック_ERR
'
    FLG_Check = False
'
    If GForm.Name = "frm_R借入一覧表" Or GForm.Name = "frm_R年度別比較表" Then
        For j = 2 To 3
            OpS(j).Value = False
            OpB(j).Value = False
            OpK(j).Value = False
            OpG(j).Value = False
            OpM(j).Value = False
            OpR(j).Value = False
        Next j
    End If
    
    If GForm.Name = "frm_R金融機関別残高表" Then
        For j = 1 To 3
            OpS(j).Value = False
            OpB(j).Value = False
            OpK(j).Value = False
            OpG(j).Value = False
            OpM(j).Value = False
            OpR(j).Value = False
        Next j
    End If
'
    If GForm.Name = "frm_R利息前払未払残高表" Then
        If OpR(0).Value = False And OpR(1).Value = False And OpR(2).Value = False And OpR(3).Value = False Then
            FLG_Check = True
        End If
    End If
'
    '杉村倉庫仕様
    If GForm.Name = "frm_R銀行別利息表" Or GForm.Name = "frm_R支払利息推移表" Then
        If OpR(0).Value = False And OpR(1).Value = False And OpR(2).Value = False And OpR(3).Value = False Then
            FLG_Check = True
        End If
    End If
'
    '集計分類
    For j = 0 To 3
        If OpR(j).Value = True And OpS(j).Value = True Then
            FLG_Check = True:   Call CEkey.SetFs(OpS(j), True) '利息区分を変えない
        ElseIf OpR(j).Value = True And OpB(j).Value = True Then
            FLG_Check = True:   Call CEkey.SetFs(OpB(j), True)
        ElseIf OpR(j).Value = True And OpK(j).Value = True Then
            FLG_Check = True:   Call CEkey.SetFs(OpK(j), True)
        ElseIf OpR(j).Value = True And OpG(j).Value = True Then
            FLG_Check = True:   Call CEkey.SetFs(OpG(j), True)
        ElseIf OpR(j).Value = True And OpM(j).Value = True Then
            FLG_Check = True:   Call CEkey.SetFs(OpM(j), True)
            
        ElseIf OpS(j).Value = True And OpB(j).Value = True Then
            FLG_Check = True:   Call CEkey.SetFs(OpS(j), True)
        ElseIf OpS(j).Value = True And OpK(j).Value = True Then
            FLG_Check = True:   Call CEkey.SetFs(OpS(j), True)
        ElseIf OpS(j).Value = True And OpG(j).Value = True Then
            FLG_Check = True:   Call CEkey.SetFs(OpS(j), True)
        ElseIf OpS(j).Value = True And OpM(j).Value = True Then
            FLG_Check = True:   Call CEkey.SetFs(OpS(j), True)
        ElseIf OpB(j).Value = True And OpK(j).Value = True Then
            FLG_Check = True:   Call CEkey.SetFs(OpB(j), True)
        ElseIf OpB(j).Value = True And OpG(j).Value = True Then
            FLG_Check = True:   Call CEkey.SetFs(OpB(j), True)
        ElseIf OpB(j).Value = True And OpM(j).Value = True Then
            FLG_Check = True:   Call CEkey.SetFs(OpB(j), True)
        ElseIf OpK(j).Value = True And OpG(j).Value = True Then
            FLG_Check = True:   Call CEkey.SetFs(OpK(j), True)
        ElseIf OpK(j).Value = True And OpM(j).Value = True Then
            FLG_Check = True:   Call CEkey.SetFs(OpK(j), True)
        ElseIf OpG(j).Value = True And OpM(j).Value = True Then
            FLG_Check = True:   Call CEkey.SetFs(OpG(j), True)
        End If
    Next j
    
    If OpG(0).Value = True And OpK(1).Value = True Then
        FLG_Check = True:   Call CEkey.SetFs(OpK(1), True)
    ElseIf OpG(0).Value = True And OpK(2).Value = True Then
        FLG_Check = True:   Call CEkey.SetFs(OpK(2), True)
    ElseIf OpG(0).Value = True And OpK(3).Value = True Then
        FLG_Check = True:   Call CEkey.SetFs(OpK(3), True)
    ElseIf OpG(1).Value = True And OpK(2).Value = True Then
        FLG_Check = True:   Call CEkey.SetFs(OpK(2), True)
    ElseIf OpG(1).Value = True And OpK(3).Value = True Then
        FLG_Check = True:   Call CEkey.SetFs(OpK(3), True)
    ElseIf OpG(2).Value = True And OpK(3).Value = True Then
        FLG_Check = True:   Call CEkey.SetFs(OpK(3), True)
    
    End If
        
    If FLG_Check = True Then
        MsgBox "指定された内容に誤りがあります"

        Exit Function
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
    Dim j As Integer
    Dim wiB_S As Integer, wiB_B As Integer, wiB_K As Integer, wiB_G As Integer
    Dim wiB_R As Integer, wiB_M As Integer
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
    '分類1～4
    For j = 0 To 3
        If OpS(j).Value = True Then
            wiB_S = j + 1
        End If
        If OpB(j).Value = True Then
            wiB_B = j + 1
        End If
        If OpK(j).Value = True Then
            wiB_K = j + 1
        End If
        If OpG(j).Value = True Then
            wiB_G = j + 1
        End If
        If OpM(j).Value = True Then
            wiB_M = j + 1
        End If
        If OpR(j).Value = True Then
            wiB_R = j + 1
        End If
    Next j
    
    '表示しない
    If OpS(4).Value = True Then
        wiB_S = 9
        CheckS = 0
    End If
    If OpB(4).Value = True Then
        wiB_B = 9
        CheckB = 0
    End If
    If OpK(4).Value = True Then
        wiB_K = 9
        CheckK = 0
    End If
    If OpG(4).Value = True Then
        wiB_G = 9
        CheckG = 0
    End If
    If OpM(4).Value = True Then
        wiB_M = 9
        CheckM = 0
    End If
    If OpR(4).Value = True Then
        wiB_R = 9
        CheckR = 0
    End If
'
    wstr = ""
    wstr = wstr & "Select * From DAKA000_帳票出力設定"
    wstr = wstr & " where フォーム名='" & GForm.Name & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.eof Then
        wRs.AddNew
        
        wRs("番号") = 90
        wRs("フォーム名") = GForm.Name
        wRs("帳票名") = P8.FCStr(L_帳票名.Caption)
    End If
    
        wRs("B_種別") = wiB_S
        wRs("B_部門") = wiB_B
        wRs("B_金融") = wiB_K
        wRs("B_銀行") = wiB_G
        wRs("B_利息") = wiB_R
        wRs("B_金利") = wiB_M
        
        wRs("P_種別") = CheckS
        wRs("P_部門") = CheckB
        wRs("P_金融") = CheckK
        wRs("P_銀行") = CheckG
        wRs("P_利息") = CheckR
        wRs("P_金利") = CheckM
        
        wRs.Update
    
    wRs.Close
    Set wRs = Nothing
'
    ' =========================================
    '           マスタ類 設定
    ' =========================================
    Call MAA070_帳票出力設定
'
    Call 画面セット
'
    ' =========================================
    '               メッセージ
    ' =========================================
    MsgBox "登録しました。", vbInformation
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
    Unload Me
    
    GForm.Enabled = True
    GForm.Show
End Sub
