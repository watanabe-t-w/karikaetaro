VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RDH010_仕訳データ 
   Caption         =   "仕訳表"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11235
   Icon            =   "RDH010_仕訳データ.dsx":0000
   StartUpPosition =   3  'Windows の既定値
   WindowState     =   2  '最大化
   _ExtentX        =   19817
   _ExtentY        =   9657
   SectionData     =   "RDH010_仕訳データ.dsx":0ECA
End
Attribute VB_Name = "RDH010_仕訳データ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "MDA010_仕訳データ"

Dim w開始年月日 As Date
Dim ps年月1 As String, ps年月2 As String
Dim w仕訳科目 As MDA010_勘定科目
Dim w推移表タイトル As MAA910_推移表タイトル
'
'------------------------------------------------
' ActiveReport_ReportStart
'------------------------------------------------
Private Sub ActiveReport_ReportStart()
'
    Dim wstr As String
    Dim FLG_TR As Boolean
'
    On Error GoTo ActiveReport_ReportStart_ERR
'
    '----------------------------------------------------------------
    '                         ** 初期設定 **
    '----------------------------------------------------------------
    'Connection
    Me.DataControl1.Connection = GDb
   
    '用紙セット
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = ddOLandscape
'
    '----------------------------------------------------------------
    '                           ** 見出し **
    '----------------------------------------------------------------
    出力日 = Now
    企業名 = GCoName
    
    If GRpt.帳票名 = "仕訳表 -月次処理-" Then
        If GRpt.テキスト_01 <> GRpt.テキスト_02 Then
            L_帳票名.Caption = GRpt.テキスト_01 & "～" & GRpt.テキスト_02 & " 仕訳表 -月次処理-"
        Else
            L_帳票名.Caption = GRpt.テキスト_01 & " 仕訳表 -月次処理-"
        End If
    
    ElseIf GRpt.帳票名 = "仕訳表 -決算処理-" Then
        L_帳票名.Caption = GRpt.テキスト_01 & " 仕訳表 -決算処理-"
    
    End If
'
    '----------------------------------------------------------------
    '          ** グループ ・ Where ・ OrderＢｙ **
    '----------------------------------------------------------------
    GroupHeader1.DataField = "GrpFld_SKUBUN" '"GrpFld_DATE"
    GroupHeader3.DataField = "I_年月"
    
    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    '** 明細ファイル 削除 **
    wstr = ""
    wstr = wstr + "Delete * From DCDA030_利息未払前払明細"
    GDb.Execute wstr
    
    wstr = ""
    wstr = wstr & "Delete * From DCDA040_仕訳データ"
    GDb.Execute wstr

    wstr = ""
    wstr = wstr & "Delete * From DCDA040_仕訳データ2"
    GDb.Execute wstr

    FLG_TR = False
    
    Call MDA010_勘定科目マスタ設定
    Call MDA010_補助科目マスタ設定
    'Call MDA010_個別補助マスタ設定
'
    '** 明細ファイル 作成 **
    '仕訳区分=1-借入金(社債)の実行
    '仕訳区分=2-借入金(社債)の返済
    '仕訳区分=3-利息(社債)の支払
    '仕訳区分=4-利息(社債)の計上
    '仕訳区分=5-社債手数料の支払
    '仕訳区分=6-社債保証料の支払
    '仕訳区分=7-借入金長期借入金長短振替

    ps年月1 = GRpt.テキスト_01
    ps年月2 = GRpt.テキスト_02
    
    If GRpt.帳票名 = "仕訳表 -月次処理-" Then
    '仕訳区分 1,2,3
        Call 月次処理
    ElseIf GRpt.帳票名 = "仕訳表 -決算処理-" Then
    '仕訳区分 4,7
        Call 決算処理
    End If
'
    '----------------------------------------------------------------
    '                          ** 合計 計算 **
    '----------------------------------------------------------------
    '----------------------------------------------------------------
    '                       ** レコード　ソース **
    '----------------------------------------------------------------
    wstr = ""
    wstr = wstr & "Select "
    wstr = wstr & "仕訳区分 & 仕訳補助 & 社債フラグ As GrpFld_SKUBUN,"
    wstr = wstr & "仕訳区分 & 仕訳補助 & 社債フラグ As I_仕訳区分,"
    wstr = wstr & "仕訳名 As I_仕訳区分名,"
    
    wstr = wstr & "Format(年月日,'" & Gfmt年月日 & "') As GrpFld_DATE,"
    'If GRpt.チェック_02 = 0 Then
    If G基本情報.日付入力区分 = "0" Then
    '和暦
        wstr = wstr & "Format(年月日,'ee年mm月dd日') As I_年月日,"
        wstr = wstr & "Format(対象年月,'ee年mm月') As I_年月,"
    Else
    '西暦
        wstr = wstr & "Format(年月日,'yyyy/mm/dd') As I_年月日,"
        wstr = wstr & "Format(対象年月,'yyyy/mm') As I_年月,"
    End If
    
    If GRpt.帳票名 = "仕訳表 -月次処理-" Then
        wstr = wstr & "番号 As I_番号,"
        wstr = wstr & "S.借入番号 As I_借入番号,"
        wstr = wstr & "G.銀行名 As I_銀行名,"
        wstr = wstr & "借方勘定科目 As I_借方科目,"
        wstr = wstr & "借方勘定科目名 As I_借方科目名,"
        wstr = wstr & "借方補助科目 As I_借方補助科目,"
        wstr = wstr & "借方補助科目名 As I_借方補助科目名,"
        wstr = wstr & "借方金額 As I_借方金額,"
        wstr = wstr & "貸方勘定科目 As I_貸方科目,"
        wstr = wstr & "貸方勘定科目名 As I_貸方科目名,"
        wstr = wstr & "貸方補助科目 As I_貸方補助科目,"
        wstr = wstr & "貸方補助科目名 As I_貸方補助科目名,"
        wstr = wstr & "貸方金額 As I_貸方金額"
        
        wstr = wstr & " FROM DCDA040_仕訳データ As S"
        wstr = wstr & " Left JOIN DAAA040_銀行マスタ As G"
        wstr = wstr & " ON S.銀行番号 = G.銀行番号"
        'wstr = wstr & " Order BY 仕訳区分,社債フラグ,仕訳補助,年月日,日番号,G.銀行番号,番号"
        'wstr = wstr & " Order BY Format(年月日,'yyyy/mm'),仕訳区分,仕訳補助,社債フラグ,年月日,S.銀行番号,S.借入番号"
        wstr = wstr & " Order BY 対象年月,仕訳区分,仕訳補助,社債フラグ,年月日,S.銀行番号,S.借入番号"
    
    ElseIf GRpt.帳票名 = "仕訳表 -決算処理-" Then
        wstr = wstr & "番号 As I_番号,"
        wstr = wstr & "S.借入番号 As I_借入番号,"
        wstr = wstr & "G.銀行名 As I_銀行名,"
        wstr = wstr & "借方勘定科目 As I_借方科目,"
        wstr = wstr & "借方勘定科目名 As I_借方科目名,"
        wstr = wstr & "借方補助科目 As I_借方補助科目,"
        wstr = wstr & "借方補助科目名 As I_借方補助科目名,"
        wstr = wstr & "借方金額 As I_借方金額,"
        wstr = wstr & "貸方勘定科目 As I_貸方科目,"
        wstr = wstr & "貸方勘定科目名 As I_貸方科目名,"
        wstr = wstr & "貸方補助科目 As I_貸方補助科目,"
        wstr = wstr & "貸方補助科目名 As I_貸方補助科目名,"
        wstr = wstr & "貸方金額 As I_貸方金額"
        
        wstr = wstr & " From DCDA040_仕訳データ As S"
        wstr = wstr & " Left JOIN DAAA040_銀行マスタ As G"
        wstr = wstr & " ON S.銀行番号 = G.銀行番号"
        'wstr = wstr & " Order BY Format(年月日,'yyyy/mm'),仕訳区分,社債フラグ,仕訳補助,年月日,日番号,S.銀行番号"
        'wstr = wstr & " Order BY Format(年月日,'yyyy/mm'),仕訳区分,仕訳補助,社債フラグ,年月日,S.銀行番号,S.借入番号"
        wstr = wstr & " Order BY 対象年月,仕訳区分,仕訳補助,社債フラグ,年月日,S.銀行番号,S.借入番号"
    End If
    
    Me.DataControl1.Source = wstr
'
    Exit Sub
'
'----------< ERROR ROUTINE >--------------------------------------------------------
ActiveReport_ReportStart_ERR:
    pERR_MES = pPROGRAM_ID + "/ ActiveReport_ReportStart() でエラー" + vbCrLf + vbCrLf + _
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
' 月次処理
'------------------------------------------------
Private Sub 月次処理()
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    
    Dim w借入データ As MAA910_借入金
'
    On Error GoTo 月次処理_ERR
'
    'tbl 仕訳データ
    '借入/支払　借入金額、元金額、利息額
    '仕訳区分 1,2,3
    
    '16/03/26 利子補給に伴う変更
    Call MBD020_借入金ワークテーブル作成("DBDA010_借入金") 'データ絞り込み
    
    wstr = ""
    wstr = wstr & "Select *"
    wstr = wstr & " From DCIA010_借入金ワーク"
    wstr = wstr & " Where 手入力区分=0"
    wstr = wstr & " And 取消フラグ=0"
    wstr = wstr & " And sm区分=0"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    Do Until wRs.eof
        w借入データ = MBD010_借入データセット(wRs)
        Call MBD010_借入金テーブル作成("", w借入データ)
        Call MDA010_仕訳現金科目作成(w借入データ)
        
        wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
'
    '仕訳区分 1,2,3,5,6
    Call MDA010_仕訳現金科目作成_明細TR("DBDA010_借入金明細TR")
'
    '利息額 銀行 金額集計 仕訳データ2→仕訳データ
    'Call MDA010_月次仕訳作成_日本ガス
'
    Exit Sub
'
'----------< ERROR ROUTINE >--------------------------------------------------------
月次処理_ERR:
    pERR_MES = pPROGRAM_ID + "/ 月次処理() でエラー" + vbCrLf + vbCrLf + _
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
' 決算処理
'------------------------------------------------
Private Sub 決算処理()
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String

    Dim j As Integer, wML As Integer, w間隔 As Integer
    Dim wNendo As Integer
    Dim wdate As Date
    Dim wSDate As String
'
    On Error GoTo 決算処理_ERR
'
    '----------< 計上　利息額 >----------
    'wデータ作成
    wdate = C年月日.平成To西暦("年月日", GRpt.テキスト_01)
    'GRpt.テキスト_01 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "e")
    '2019/01/15 日付入力区分 仕様変更
    If G基本情報.日付入力区分 = "0" Then
    '和暦
        GRpt.テキスト_01 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "e")
    Else
    '西暦
        GRpt.テキスト_01 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "y")
    End If
    
    GRpt.推移 = "月次"
    
    '16/03/26 利子補給に伴う変更
    Call MBD020_借入金ワークテーブル作成("DBDA010_借入金") 'データ絞り込み
    
'    Call MRB010_標準入力借入残高表固定日数("DBDA010_借入金")
'    Call MRB010_手入力借入残高表("DBDA010_借入金")
    '16/03/26 利子補給に伴う変更
    Call MRB010_標準入力借入残高表固定日数("DCIA010_借入金ワーク")
    Call MRB010_手入力借入残高表("DCIA010_借入金ワーク")
    '
    '仕訳伝票作成
    GInt1 = 0
    'w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "平成")
    '2019/01/15 日付入力区分 仕様変更
    If G基本情報.日付入力区分 = "0" Then
    '和暦
        If Len(GRpt.テキスト_01) <= 2 Then
            w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "平成")
        Else
        w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "西暦")
        End If
    Else
    '西暦
        w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "西暦")
    End If
    
    w推移表タイトル = MUA010_推移表ファイル作成("", "", w開始年月日, GRpt.推移, 12)
    For j = 1 To 12
        If ps年月1 = w推移表タイトル.X番目年月(j) Then
                GInt1 = j
            Exit For
        End If
    Next
    
    If GInt1 < 1 Then
        Exit Sub
    End If
    
    GRpt.テキスト_01 = ps年月1
    GRpt.テキスト_02 = ps年月2
    Call MDA010_仕訳計上科目作成_残高
'

'
    '----------< 元金額 >----------
    'wデータ作成
    Select Case G基本情報.決算サイクル
    Case 1
    '月次決算
        '仮の決算月を指定付に設定し、G基本情報.決算サイクル=年次と同処理をする
        wNendo = G基本情報.決算月
        G基本情報.決算月 = CInt(Format(GRpt.テキスト_01, "mm"))
        
        GRpt.推移 = "年次":   wML = 10: w間隔 = 12
    Case 3
        GRpt.推移 = "四半期": wML = 12: w間隔 = G基本情報.決算サイクル
    Case 6
        GRpt.推移 = "半期":   wML = 10: w間隔 = G基本情報.決算サイクル
    Case Else
        GRpt.推移 = "年次":   wML = 10: w間隔 = 12
    End Select
    
    wdate = C年月日.平成To西暦("年月日", GRpt.テキスト_01)
    'GRpt.テキスト_01 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "e")
    '2019/01/15 日付入力区分 仕様変更
    If G基本情報.日付入力区分 = "0" Then
    '和暦
        GRpt.テキスト_01 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "e")
    Else
    '西暦
        GRpt.テキスト_01 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "y")
    End If
    
'    Call MRB010_標準入力借入残高表固定日数("DBDA010_借入金")
'    Call MRB010_手入力借入残高表("DBDA010_借入金")
    '16/03/26 利子補給に伴う変更
    Call MRB010_標準入力借入残高表("DCIA010_借入金ワーク")
    Call MRB010_手入力借入残高表("DCIA010_借入金ワーク")
    '
    '仕訳伝票作成
    GInt1 = 0
    'w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "平成")
    '2019/01/15 日付入力区分 仕様変更
    If G基本情報.日付入力区分 = "0" Then
    '和暦
        If Len(GRpt.テキスト_01) <= 2 Then
            w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "平成")
        Else
        w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "西暦")
        End If
    Else
    '西暦
        w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "西暦")
    End If
    w推移表タイトル = MUA010_推移表ファイル作成("", "", w開始年月日, GRpt.推移, wML)
    
    '翌回から1年間
    If G基本情報.日付入力区分 = 0 Then
        wdate = DateAdd("m", w間隔, C年月日.平成To西暦("年月", ps年月1))
    Else
        wdate = DateAdd("m", w間隔, ps年月1)
    End If
    wSDate = Format(wdate, Gfmt年月)
    
    For j = 1 To wML
        If wSDate = w推移表タイトル.X番目年月(j) Then
                GInt1 = j
            Exit For
        End If
    Next
    
    If GInt1 < 1 Then
        Exit Sub
    End If
    
    GRpt.テキスト_01 = ps年月1
    GRpt.テキスト_02 = ps年月2
    Call MDA010_仕訳長短振替科目作成
    '
    If G基本情報.決算サイクル = 1 Then
    '決算月を元に戻す
        G基本情報.決算月 = wNendo
    End If
'
    Exit Sub
'
'----------< ERROR ROUTINE >--------------------------------------------------------
決算処理_ERR:
    pERR_MES = pPROGRAM_ID + "/ 決算処理() でエラー" + vbCrLf + vbCrLf + _
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
' ActiveReport_ReportEnd
'------------------------------------------------
Private Sub ActiveReport_ReportEnd()
'
    'FBA010_帳票範囲指定.メッセージ = ""
    'FBA010_帳票範囲指定.メッセージ.Refresh
'
    ' =========================================
    '           　 CsvFile 作成
    ' =========================================
    If GRpt.CSV = 1 Then
        Call MX040_CsvOut_KARI
    End If
'
    ' =========================================
    '           　 ボタン制御
    ' =========================================
    'FBA010_帳票範囲指定.実行.Enabled = True
    'FBA010_帳票範囲指定.閉じる.Enabled = True
'
    'FBA010_帳票範囲指定.拡張.SetFocus
'
'
    ' =========================================
    '  借換たろう！お試し版帳票出力回数チェック
    ' =========================================
    If GSys.Sys = "借入金 お試し版" Then
        Call MAA001_KARIKAETAROU_CNT
    End If
'
End Sub

'------------------------------------------------
' ActiveReport_NoData
'------------------------------------------------
Private Sub ActiveReport_NoData()
'
    'FBA010_帳票範囲指定.メッセージ = "出力すべきデータはありません"
    'FBA010_帳票範囲指定.メッセージ.Refresh
    GSstrt帳票Msg = "出力すべきデータはありません"
'
    Me.Cancel
    DoEvents
'
    ' =========================================
    '           　 ボタン制御
    ' =========================================
    'FBA010_帳票範囲指定.実行.Enabled = True
    'FBA010_帳票範囲指定.閉じる.Enabled = True
'
    'FBA010_帳票範囲指定.拡張.SetFocus
'
    Unload Me
'
End Sub

'------------------------------------------------
' ActiveReport_Error
'------------------------------------------------
Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
'
    'FBA010_帳票範囲指定.メッセージ = "出力できませんでした"
    'FBA010_帳票範囲指定.メッセージ.Refresh
    GSstrt帳票Msg = "出力できませんでした"
'
    Me.Cancel
    DoEvents

    ' =========================================
    '           　 ボタン制御
    ' =========================================
    'FBA010_帳票範囲指定.実行.Enabled = True
    'FBA010_帳票範囲指定.閉じる.Enabled = True
'
    'FBA010_帳票範囲指定.拡張.SetFocus
'
    Unload Me
'
End Sub
