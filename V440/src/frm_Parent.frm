VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frm_Parent 
   BackColor       =   &H8000000C&
   Caption         =   "借換たろうSP"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   12615
   Icon            =   "frm_Parent.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuファイル 
      Caption         =   "ファイル"
      Begin VB.Menu mnuログアウト 
         Caption         =   "ログアウト"
      End
      Begin VB.Menu mnu終了 
         Caption         =   "終了"
      End
   End
   Begin VB.Menu mnu基本設定 
      Caption         =   "基本設定"
      Begin VB.Menu mnu固定項目登録 
         Caption         =   "固定項目登録"
      End
      Begin VB.Menu mnu祝日マスタ 
         Caption         =   "祝日マスタ"
      End
   End
   Begin VB.Menu mnuマスタ登録 
      Caption         =   "マスタ登録"
      Begin VB.Menu mnu銀行マスタ 
         Caption         =   "銀行マスタ"
      End
      Begin VB.Menu mnu借入種別マスタ 
         Caption         =   "借入種別マスタ"
      End
      Begin VB.Menu mnu基準金利マスタ 
         Caption         =   "基準金利マスタ"
      End
      Begin VB.Menu mnu金利シミュレーションGPマスタ 
         Caption         =   "金利ｼﾐｭﾚｰｼｮﾝGPマスタ"
      End
      Begin VB.Menu mnu部門マスタ 
         Caption         =   "部門マスタ"
      End
   End
   Begin VB.Menu mnu借入金登録 
      Caption         =   "借入金登録"
   End
   Begin VB.Menu mnu帳票出力 
      Caption         =   "帳票出力"
      Begin VB.Menu mnu借入金台帳 
         Caption         =   "借入金台帳"
      End
      Begin VB.Menu mnu借入明細表 
         Caption         =   "借入明細表"
      End
      Begin VB.Menu mnu返済予定表 
         Caption         =   "返済予定表"
      End
      Begin VB.Menu mnu借入一覧表 
         Caption         =   "借入一覧表"
      End
      Begin VB.Menu mnu借入残高表 
         Caption         =   "借入残高表"
      End
      Begin VB.Menu mnu残高推移表 
         Caption         =   "残高推移表"
      End
      Begin VB.Menu mnu利息前払未払明細表 
         Caption         =   "利息明細表"
      End
      Begin VB.Menu mnu利息前払未払残高表 
         Caption         =   "利息残高表"
      End
      Begin VB.Menu mnu利息残高推移表 
         Caption         =   "利息残高推移表"
      End
      Begin VB.Menu mnu損益利息一覧表 
         Caption         =   "損益利息一覧表"
      End
      Begin VB.Menu mnu平均金利平均残高表 
         Caption         =   "平均金利平均残高表"
      End
      Begin VB.Menu mnu平均金利平均残高推移表 
         Caption         =   "平均金利平均残高推移表"
      End
      Begin VB.Menu mnu金融機関別残高表 
         Caption         =   "金融機関別残高表"
      End
      Begin VB.Menu mnu年度別残高表 
         Caption         =   "年度別比較表"
      End
   End
   Begin VB.Menu mnu金利シミュレーション 
      Caption         =   "金利シミュレーション"
      Begin VB.Menu mnu金利シミュレーション入力 
         Caption         =   "金利シミュレーション入力"
      End
      Begin VB.Menu mnu金利SM借入明細表 
         Caption         =   "借入明細表"
      End
      Begin VB.Menu mnu金利SM借入残高表 
         Caption         =   "借入残高表"
      End
      Begin VB.Menu mnu金利SM借入残高推移表 
         Caption         =   "借入残高推移表"
      End
      Begin VB.Menu mnu金利SM利息前払未払残高表 
         Caption         =   "利息残高表"
      End
      Begin VB.Menu mnu金利SM利息残高推移表 
         Caption         =   "利息残高推移表"
      End
   End
   Begin VB.Menu mnu仕訳マスタ登録 
      Caption         =   "仕訳マスタ登録"
      Begin VB.Menu mnu仕訳データ作成 
         Caption         =   "月次仕訳データ作成"
      End
      Begin VB.Menu mnu決算仕訳データ作成 
         Caption         =   "決算仕訳データ作成"
      End
      Begin VB.Menu mnu勘定科目マスタ 
         Caption         =   "勘定科目マスタ"
      End
      Begin VB.Menu mnu補助科目マスタ 
         Caption         =   "補助科目マスタ"
      End
   End
   Begin VB.Menu mnu時価評価 
      Caption         =   "時価評価"
      Begin VB.Menu mnu基準金利レート 
         Caption         =   "基準金利レート"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu借入金時価評価明細表 
         Caption         =   "借入金時価評価明細表"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu借入金時価評価適用金利一覧 
         Caption         =   "借入金時価評価適用金利一覧"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu借入金時価評価一覧表 
         Caption         =   "借入金時価評価一覧表"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnu運用管理 
      Caption         =   "運用管理"
      Begin VB.Menu mnuユーザー設定 
         Caption         =   "ユーザー設定"
      End
      Begin VB.Menu mnuログ照会 
         Caption         =   "ログ照会"
      End
   End
End
Attribute VB_Name = "frm_Parent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Option Explicit
'
Private Const pPROGRAM_ID As String = "メインフォーム"
'
Dim wiTab As Integer
Dim wsBut As String

Dim wdate As Date

Dim wRs2 As ADODB.Recordset
Dim wstr As String

Dim CopyFLG As Integer    '複製判断フラグ　false:新規　true:複製
Dim CopyDB名  As String   '複製対象DB名

Dim FLG_KIGYO As Boolean, FLG_DIR As Boolean, FLG_LOSTK As Boolean, FLG_New As Boolean
Dim wi_Kosin1 As Integer, wi_Kosin2 As Integer
Dim wl_st As Long
Dim wDB名 As String, wBK As String
Dim w企業名Key As String, wFlg企業名Key As String, wNew企業名Key As String

Dim wi決算月 As Integer, wi決算締日 As Integer, wi回収有無 As Integer, wi支払有無 As Integer, wd実績年月 As Date
Dim wslog As String
'----------< Msg >------------------------------------------------------------------
Private Const Msg_01 = "しばらくしてから処理を行ってください"

Private Sub MDIForm_Activate()
'------------------------------------------------
' Form_Activate
'------------------------------------------------
    GStr = "": GStr_1 = "": GStr_2 = ""
    GInt1 = 0: GInt2 = 0
End Sub

Private Sub MDIForm_Initialize()
'
'------------------------------------------------
' Form_Initialize
'------------------------------------------------
    wsBut = ""
End Sub

Private Sub MDIForm_Load()
'------------------------------------------------
' Form_Load
'------------------------------------------------
'
    Dim j As Integer
'
    DoEvents
    
    GInputDateKbn = "1"
    
    ' =========================================
    '                 初期設定
    ' =========================================
    frm_Fメインフォーム.Show
    frm_Fメインフォーム.Left = 0
    frm_Fメインフォーム.Top = 0
    
    G会議 = ""
    ReDim G独算(0)

    '----------< mdb Open >---------------------------------------------------------
    Call AdoDbOpen("Jet", GDb, GDbName, "", , GPwd)
'
    wdate = GDate1
    Call Set_VerNo
    Call Set_CoName
'
    If GSys.Mem = "複数" Then
        GFcap = GVerNo + Space(1) + GCoName
    Else
        GFcap = GVerNo
    End If
    Me.Caption = GFcap
    
    'Me.Caption = Me.Caption & " " & GCoName
'
    ' =========================================
    '           基本 マスタ等 設定
    ' =========================================
    'Call MRD010_マスタ_Read
    
    ' =========================================
    '           基本情報ファイル Read
    ' =========================================
    MAA010_基本情報ファイル_Read
    MAA020_コントロールファイル_Read
    
    ' =========================================
    '           マスタ類 設定
    ' =========================================
    Call MAA030_銀行マスタ設定
    Call MAA050_保証率マスタ設定
    Call MAA060_税率マスタ設定
    
    '2017/12/01 祝日マスタ ADD
    Call MAA090_祝日マスタ設定

    Call MAA070_借入金種別設定
    Call MAA070_金利グループ設定
    Call MAA070_金利SM率設定
    Call MAA070_基準金利設定
    Call MAA070_部門設定
    Call MAA070_帳票出力設定
    
    Call MAA200_基準金利レート設定

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'
    ReDim G科目マスタ(0)
    ReDim G償却率マスタ(0)
    ReDim G税率マスタ(0)

    Call MXA030_MCLEAR
'
    DoEvents

    GDb.Close
    Set GDb = Nothing
'
    wsBut = ""

    If G実績共有 = "共有" Then
        Call DEL_JISEKIL_TBL
    End If
'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    Call MXA030_LOG_WRITE("ログアウト", "ログアウト", "")
'
    If GSys.Mem = "単一" Then
        End
    Else
        Unload Me
    End If

    frm_Fデータベース選択.Show

End Sub

Private Sub mnuユーザー設定_Click()

    ' =========================================
    '           権限チェック
    ' =========================================
    Select Case GUserKen
        Case "0"
            '入力権限
            MsgBox "権限がありません", vbExclamation
            Exit Sub
        Case "1"
            '照会権限
            MsgBox "権限がありません", vbExclamation
            Exit Sub
        Case "5"
            '管理者権限
            frm_Mユーザー設定.Show
        Case Else
            MsgBox "権限がありません", vbExclamation
            Exit Sub
    End Select

End Sub

Private Sub mnuログアウト_Click()
    Unload Me
End Sub

Private Sub mnuログ照会_Click()
    frm_Fログ照会.Show
End Sub

Private Sub mnu借入種別マスタ_Click()
    UNLOAD_MASFRM
    frm_M借入種別マスタ.Show
End Sub

Private Sub mnu基準金利マスタ_Click()
    UNLOAD_MASFRM
    frm_M基準金利マスタ.Show
End Sub

Private Sub mnu金利シミュレーションGPマスタ_Click()

    UNLOAD_MASFRM
    frm_M金利シミュレーショングループマスタ.Show
    
End Sub

Private Sub mnu銀行マスタ_Click()
    UNLOAD_MASFRM
    frm_M銀行マスタ.Show
End Sub

Private Sub mnu部門マスタ_Click()
    UNLOAD_MASFRM
    frm_M部門マスタ.Show
End Sub

Private Sub mnu固定項目登録_Click()
    frm_I固定項目登録.Show
End Sub

Private Sub mnu祝日マスタ_Click()
    frm_M祝日マスタ.Show
End Sub

Private Sub mnu借入金台帳_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R借入金台帳.Show
    
End Sub

Private Sub mnu借入金登録_Click()
    frm_I借入金登録.Show
End Sub

Private Sub mnu終了_Click()
    Unload Me
    
    End
End Sub

Private Sub mnu借入明細表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R借入金明細表.Show
End Sub

Private Sub mnu借入一覧表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R借入一覧表.Show
End Sub

Private Sub mnu残高一覧表_Click()
    GStr = ""
    frm_R借入残高表.Show
End Sub

Private Sub mnu返済予定表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R返済予定表.Show
End Sub

Private Sub mnu借入残高表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R借入残高表.Show
End Sub

Private Sub mnu残高推移表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R借入残高推移表.Show
End Sub

Private Sub mnu借入残高推移表_Click()
    GStr = ""
    frm_R借入残高推移表.Show
End Sub

Private Sub mnu利息残高推移表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R利息残高推移表.Show
End Sub

Private Sub mnu借入利息残高推移表_Click()
    GStr = ""
    frm_R利息残高推移表.Show
End Sub

Private Sub mnu損益利息一覧表_Click()
    GStr = ""
    frm_R損益利息一覧表.Show
End Sub

Private Sub mnu金融機関別残高表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R金融機関別残高表.Show
End Sub

Private Sub mnu年度別残高表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R年度別比較表.Show
End Sub

'Private Sub mnu簡易資金繰表_Click()
'    UNLOAD_REPFRM
'    GStr = ""
'    frm_R簡易資金繰表.Show
'End Sub

Private Sub mnu金利シミュレーション入力_Click()
    frm_I金利シミュレーション入力.Show
End Sub

Private Sub mnu金利SM借入明細表_Click()
    UNLOAD_REPFRM
    GStr = "金利GR"
    frm_R借入金明細表.Show
End Sub

Private Sub mnu金利SM借入残高表_Click()
    UNLOAD_REPFRM
    GStr = "金利GR"
    frm_R借入残高表.Show
End Sub

Private Sub mnu金利SM借入残高推移表_Click()
    UNLOAD_REPFRM
    GStr = "金利GR"
    frm_R借入残高推移表.Show
End Sub

Private Sub mnu金利SM利息前払未払残高表_Click()
    UNLOAD_REPFRM
    GStr = "金利GR"
    frm_R利息前払未払残高表.Show
End Sub

Private Sub mnu金利SM利息残高推移表_Click()
    UNLOAD_REPFRM
    GStr = "金利GR"
    frm_R利息残高推移表.Show
End Sub

Private Sub mnu勘定科目マスタ_Click()
    UNLOAD_MASFRM
    frm_M勘定科目マスタ.Show
End Sub

Private Sub mnu補助科目マスタ_Click()
    UNLOAD_MASFRM
    frm_M補助科目マスタ.Show
End Sub

Private Sub mnu個別補助科目マスタ_Click()
    UNLOAD_MASFRM
    frm_M個別補助科目マスタ.Show
End Sub

Private Sub mnu利息前払未払明細表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R利息前払未払明細表.Show
End Sub

Private Sub mnu利息前払未払残高表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R利息前払未払残高表.Show
End Sub

Private Sub mnu仕訳データ作成_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R仕訳データ作成.Show
End Sub

Private Sub mnu決算仕訳データ作成_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R決算仕訳データ作成.Show
End Sub

Private Sub mnu平均金利平均残高推移表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R平均金利平均残高推移表.Show
End Sub

Private Sub mnu平均金利平均残高表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R平均金利平均残高表.Show
End Sub

Private Sub mnu基準金利レート_Click()
    UNLOAD_MASFRM
    frm_M長期プライムレート.Show
End Sub

Private Sub mnu借入金時価評価明細表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R借入金時価評価明細表.Show
End Sub

Private Sub mnu借入金時価評価適用金利一覧_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R借入金時価評価適用金利一覧.Show
End Sub

Private Sub mnu借入金時価評価一覧表_Click()
    UNLOAD_REPFRM
    GStr = ""
    frm_R借入金時価評価一覧表.Show
End Sub

''杉村倉庫仕様
'Private Sub mnu1年内返済集計表_Click()
'    UNLOAD_REPFRM
'    GStr = ""
'    frm_R1年内返済集計表.Show
'End Sub
'
'Private Sub mnu銀行別利息表_Click()
'    UNLOAD_REPFRM
'    GStr = ""
'    frm_R銀行別利息表.Show
'End Sub
'
'Private Sub mnu支払利息推移表_Click()
'    UNLOAD_REPFRM
'    GStr = ""
'    frm_R支払利息推移表.Show
'End Sub

'------------------------------------------------
' Set_CoName
'------------------------------------------------
Private Sub Set_CoName()
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
'
    wstr = ""
    wstr = wstr + "Select * From DAAA070_企業名マスタ"
'    wstr = wstr + " Where System = 'System'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
        GKeyName = P8.FCStr(wRs("企業名Key"))
        GCoName = P8.FCStr(wRs("企業名"))
    End If
    wRs.Close
    Set wRs = Nothing
'
    If GCoName = "---------------" Then
        GCoName = ""
    End If
'
End Sub

'------------------------------------------------
' Set_VerNo
'------------------------------------------------
Private Sub Set_VerNo()
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
'
    wstr = ""
    wstr = wstr + "Select * From DAAA000_バージョン"
    wstr = wstr + " Where System = 'System'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
        GVerNo = P8.FCStr(wRs("Version"))
    End If
    wRs.Close
    Set wRs = Nothing
'
End Sub

'------------------------------------------------
' バックアップ
'------------------------------------------------
Private Sub バックアップ()
'
'
'----------< ERROR ROUTINE >---------------------------------------------------
バックアップ_ERR:
    pERR_MES = pPROGRAM_ID + "/ バックアップ() でエラー" + vbCrLf + vbCrLf + _
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
' DELETE_NEWREC
'------------------------------------------------
Private Sub DELETE_NEWREC()
'
    On Error GoTo DELETE_NEWREC_ERR
'
    wstr = ""
    wstr = wstr + "Delete"
    wstr = wstr + " From DAAA070_企業名マスタ"
    wstr = wstr + " Where 企業名Key = '" + wNew企業名Key + "'"
    GDb2.Execute wstr
            
    wNew企業名Key = ""
    
    DoEvents
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
DELETE_NEWREC_ERR:
    pERR_MES = pPROGRAM_ID + "/ DELETE_NEWREC() でエラー" + vbCrLf + vbCrLf + _
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
' DEL_JISEKIL_TBL
'------------------------------------------------
Private Sub DEL_JISEKIL_TBL()
'
    '----------< Client KXXX.mdb Open >---------------------------------------------
    GRet = ZMA020_OpenDatabase(DB_Client, GDbName, False, GPwd)
    
    '----------< Delete DBBA010_売上実績L >-----------------------------------------
    Call ZMA020_TableDel(DB_Client, "DBBA010_売上実績L")

    '----------< Client KXXX.mdb Close >--------------------------------------------
    DB_Client.Close
'
End Sub



