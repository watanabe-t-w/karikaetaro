VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_Fメインフォーム 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'なし
   Caption         =   "&H80000000&"
   ClientHeight    =   10965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2325
   Icon            =   "frm_Fメインフォーム.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10965
   ScaleWidth      =   2325
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "登録"
      TabPicture(0)   =   "frm_Fメインフォーム.frx":0ECA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(5)=   "L_T回数"
      Tab(0).Control(6)=   "L_回数"
      Tab(0).Control(7)=   "勘定科目マスタ"
      Tab(0).Control(8)=   "固定項目登録"
      Tab(0).Control(9)=   "借入金登録"
      Tab(0).Control(10)=   "補助科目マスタ"
      Tab(0).Control(11)=   "部門登録"
      Tab(0).Control(12)=   "祝日マスタ"
      Tab(0).Control(13)=   "銀行マスタ"
      Tab(0).Control(14)=   "借入種別マスタ"
      Tab(0).Control(15)=   "基準金利マスタ"
      Tab(0).Control(16)=   "金利シミュレーションGP"
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "帳票"
      TabPicture(1)   =   "frm_Fメインフォーム.frx":0EE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label12"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "仕訳データ作成"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "借入金台帳"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "借入明細表"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "返済予定表"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "借入一覧表"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "残高一覧表"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "残高推移表"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "利息前払未払明細表"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "利息前払未払残高表"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "利息残高推移表"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "平均金利平均残高推移表"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "年度別残高表"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "決算仕訳データ作成"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "損益利息一覧表"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "金融機関別残高表"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "平均金利平均残高表"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).ControlCount=   18
      Begin VB.CommandButton 金利シミュレーションGP 
         Caption         =   "金利ｼﾐｭﾚｰｼｮﾝGP"
         Height          =   375
         Left            =   -74880
         TabIndex        =   4
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton 基準金利マスタ 
         Caption         =   "基準金利マスタ"
         Height          =   375
         Left            =   -74880
         TabIndex        =   3
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton 借入種別マスタ 
         Caption         =   "借入種別マスタ"
         Height          =   375
         Left            =   -74880
         TabIndex        =   2
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton 銀行マスタ 
         Caption         =   "銀行マスタ"
         Height          =   375
         Left            =   -74880
         TabIndex        =   1
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CommandButton 祝日マスタ 
         Caption         =   "祝日マスタ"
         Height          =   375
         Left            =   -74880
         TabIndex        =   9
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton 平均金利平均残高表 
         Caption         =   "平均金利平均残高表"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   3960
         Width           =   1815
      End
      Begin VB.CommandButton 金融機関別残高表 
         Caption         =   "金融機関別残高表"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   3600
         Width           =   1815
      End
      Begin VB.CommandButton 損益利息一覧表 
         Caption         =   "損益利息一覧表"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CommandButton 決算仕訳データ作成 
         Caption         =   "決算仕訳データ作成"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   6720
         Width           =   1815
      End
      Begin VB.CommandButton 部門登録 
         Caption         =   "部門登録"
         Height          =   375
         Left            =   -74880
         TabIndex        =   5
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CommandButton 年度別残高表 
         Caption         =   "年度別比較表"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   5520
         Width           =   1815
      End
      Begin VB.CommandButton 平均金利平均残高推移表 
         Caption         =   "平均金利平残推移表"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   5160
         Width           =   1815
      End
      Begin VB.CommandButton 利息残高推移表 
         Caption         =   "利息残高推移表"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   4800
         Width           =   1815
      End
      Begin VB.CommandButton 利息前払未払残高表 
         Caption         =   "利息残高表"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2880
         Width           =   1815
      End
      Begin VB.CommandButton 利息前払未払明細表 
         Caption         =   "利息明細表"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CommandButton 残高推移表 
         Caption         =   "残高推移表"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   4440
         Width           =   1815
      End
      Begin VB.CommandButton 残高一覧表 
         Caption         =   "借入残高表"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CommandButton 借入一覧表 
         Caption         =   "借入一覧表"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CommandButton 返済予定表 
         Caption         =   "返済予定表"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton 借入明細表 
         Caption         =   "借入明細表"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton 借入金台帳 
         Caption         =   "借入金台帳"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton 仕訳データ作成 
         Caption         =   "月次仕訳データ作成"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   6360
         Width           =   1815
      End
      Begin VB.CommandButton 補助科目マスタ 
         Caption         =   "補助科目マスタ"
         Height          =   375
         Left            =   -74880
         TabIndex        =   7
         Top             =   5040
         Width           =   1815
      End
      Begin VB.CommandButton 借入金登録 
         Caption         =   "借入金登録"
         Height          =   375
         Left            =   -74880
         TabIndex        =   10
         Top             =   5880
         Width           =   1815
      End
      Begin VB.CommandButton 固定項目登録 
         Caption         =   "固定項目登録"
         Height          =   375
         Left            =   -74880
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton 勘定科目マスタ 
         Caption         =   "勘定科目マスタ"
         Height          =   375
         Left            =   -74880
         TabIndex        =   6
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label L_回数 
         Alignment       =   1  '右揃え
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   7920
         Width           =   1575
      End
      Begin VB.Label L_T回数 
         Height          =   255
         Left            =   -74880
         TabIndex        =   37
         Top             =   7560
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "部門登録-----------"
         Height          =   255
         Left            =   -74880
         TabIndex        =   35
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "仕訳登録-----------"
         Height          =   255
         Left            =   -74880
         TabIndex        =   34
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "借入金マスタ登録----"
         Height          =   255
         Left            =   -74880
         TabIndex        =   33
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "仕訳データ----------"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   6120
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "基本設定-----------"
         Height          =   255
         Left            =   -74880
         TabIndex        =   31
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "借入金管理---------"
         Height          =   255
         Left            =   -74880
         TabIndex        =   30
         Top             =   5640
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "帳票出力-----------"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.CommandButton 終了 
      Caption         =   "終了"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   9720
      Width           =   2055
   End
   Begin 借換たろう.ZU050_Button ZU050_Button1 
      Height          =   375
      Left            =   120
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BackColor       =   16777215
      BorderColor     =   8421504
      Shape           =   4
      ForeColor       =   33023
      Caption         =   "メニュー"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frm_Fメインフォーム"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''------------------------------------------------
'' Form_Initialize
''------------------------------------------------
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
' Form_KeyPress
'------------------------------------------------
Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = CEkey.X020_EnterKey(Me, KeyAscii, True)
End Sub

Private Sub Form_Load()
'
    frm_Fメインフォーム.L_T回数.Caption = ""
    frm_Fメインフォーム.L_回数.Caption = ""
    If GSys.Sys = "借入金 お試し版" Then
        Call MAA001_KARIKAETAROU_PRE
        frm_Fメインフォーム.L_T回数.Caption = "お試し版"
    End If
'
End Sub

'------------------------------------------------
'基本設定
'------------------------------------------------
Private Sub 固定項目登録_Click()
    frm_I固定項目登録.Show
End Sub

Private Sub 祝日マスタ_Click()
    frm_M祝日マスタ.Show
End Sub

'------------------------------------------------
'マスタ登録
'------------------------------------------------
Private Sub 基準金利マスタ_Click()
    UNLOAD_MASFRM
    frm_M基準金利マスタ.Show
End Sub

Private Sub 金利シミュレーションGP_Click()
    UNLOAD_MASFRM
    frm_M金利シミュレーショングループマスタ.Show
End Sub

Private Sub 銀行マスタ_Click()
    UNLOAD_MASFRM
    frm_M銀行マスタ.Show
End Sub

Private Sub 借入種別マスタ_Click()
    UNLOAD_MASFRM
    frm_M借入種別マスタ.Show
End Sub

Private Sub 部門登録_Click()
    frm_M部門マスタ.Show
End Sub

Private Sub 長期プライムレート_Click()
    UNLOAD_MASFRM
    frm_M長期プライムレート.Show
End Sub

'------------------------------------------------
'借入金管理
'------------------------------------------------
Private Sub 借入金登録_Click()
    frm_I借入金登録.Show
End Sub

'------------------------------------------------
'帳票出力
'------------------------------------------------
Private Sub 借入一覧表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R借入一覧表.Show
End Sub

Private Sub 借入金台帳_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R借入金台帳.Show
End Sub

Private Sub 借入明細表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R借入金明細表.Show
End Sub

Private Sub 返済予定表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R返済予定表.Show
End Sub

Private Sub 残高一覧表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R借入残高表.Show
End Sub

Private Sub 残高推移表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R借入残高推移表.Show
End Sub

Private Sub 利息前払未払明細表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R利息前払未払明細表.Show
End Sub

Private Sub 利息前払未払残高表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R利息前払未払残高表.Show
End Sub

Private Sub 損益利息一覧表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R損益利息一覧表.Show
End Sub

Private Sub 利息残高推移表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R利息残高推移表.Show
End Sub

Private Sub 平均金利平均残高表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R平均金利平均残高表.Show
End Sub

Private Sub 平均金利平均残高推移表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R平均金利平均残高推移表.Show
End Sub

Private Sub 年度別残高表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R年度別比較表.Show
End Sub

Private Sub 金融機関別残高表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R金融機関別残高表.Show
End Sub

'Private Sub 簡易資金繰表_Click()
'    UNLOAD_REPFRM
'    frm_R簡易資金繰表.Show
'End Sub

'------------------------------------------------
'帳票出力　時価評価
'------------------------------------------------
Private Sub 借入金時価評価一覧表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R借入金時価評価一覧表.Show
End Sub

Private Sub 借入金時価評価適用金利一覧_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R借入金時価評価適用金利一覧.Show
End Sub

Private Sub 借入金時価評価明細表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R借入金時価評価明細表.Show
End Sub

'------------------------------------------------
'帳票出力　金利シミュレーション
'------------------------------------------------
'Private Sub 金利SM借入明細表_Click()
'    UNLOAD_REPFRM
'    GStr = "金利GR"
'    frm_R借入金明細表.Show
'End Sub
'
'Private Sub 金利SM借入残高表_Click()
'    UNLOAD_REPFRM
'    GStr = "金利GR"
'    frm_R借入残高表.Show
'End Sub
'
'Private Sub 金利SM借入残高推移表_Click()
'    UNLOAD_REPFRM
'    GStr = "金利GR"
'    frm_R借入残高推移表.Show
'End Sub
'
'Private Sub 金利SM利息前払未払残高表_Click()
'    UNLOAD_REPFRM
'    GStr = "金利GR"
'    frm_R利息前払未払残高表.Show
'End Sub
'
'Private Sub 金利SM利息残高推移表_Click()
'    UNLOAD_REPFRM
'    GStr = "金利GR"
'    frm_R利息残高推移表.Show
'End Sub

'------------------------------------------------
'シミュレーション
'------------------------------------------------
'Private Sub 金利シミュレーション入力_Click()
'    frm_I金利シミュレーション入力.Show
'End Sub

'------------------------------------------------
'仕訳
'------------------------------------------------
Private Sub 勘定科目マスタ_Click()
    UNLOAD_MASFRM
    frm_M勘定科目マスタ.Show
End Sub

Private Sub 補助科目マスタ_Click()
    UNLOAD_MASFRM
    frm_M補助科目マスタ.Show
End Sub

'日本ガス
'Private Sub 個別補助科目マスタ_Click()
'    UNLOAD_MASFRM
'    frm_M個別補助科目マスタ.Show
'End Sub

Private Sub 仕訳データ作成_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R仕訳データ作成.Show
End Sub

Private Sub 決算仕訳データ作成_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R決算仕訳データ作成.Show
End Sub

''杉村倉庫仕様
'Private Sub 長短振替集計表_Click()
'    UNLOAD_REPFRM
'    GStr = ""
'    frm_R1年内返済集計表.Show
'End Sub
'
''杉村倉庫仕様
'Private Sub 銀行別利息表_Click()
'    UNLOAD_REPFRM
'    GStr = ""
'    frm_R銀行別利息表.Show
'End Sub
'
''杉村倉庫仕様
'Private Sub 支払利息推移表_Click()
'    UNLOAD_REPFRM
'    GStr = ""
'    frm_R支払利息推移表.Show
'End Sub

Private Sub 終了_Click()
    Unload Me
    Unload frm_Parent
    
    End
End Sub


