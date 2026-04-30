VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RDG010_残高表グラフ 
   Caption         =   "金融機関別残高表"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10995
   Icon            =   "RDG010_残高表グラフ.dsx":0000
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   _ExtentX        =   19394
   _ExtentY        =   12806
   SectionData     =   "RDG010_残高表グラフ.dsx":0ECA
End
Attribute VB_Name = "RDG010_残高表グラフ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "RDG010_残高表グラフ"
'
Dim wRs As ADODB.Recordset

Dim wstr As String
Dim wWhere As String

Dim w分母 As Integer
Dim wiCnt As Integer

'グループ集計
Dim wiGCnt As Integer
Dim ws_GrpName() As String
Dim wdYusi() As Double, wdTokiYusi() As Double, wdzan() As Double
Dim wdGankin() As Double, wdRisoku() As Double, wdHensai() As Double
Dim wdPYusi() As Double, wdPTokiYusi() As Double, wdPZan() As Double
Dim wdPGankin() As Double, wdPRisoku() As Double, wdPHensai() As Double

Dim pColor(9) As String

Dim w推移表タイトル As MAA910_推移表タイトル
'
'------------------------------------------------
' ActiveReport_ReportStart
'------------------------------------------------
Private Sub ActiveReport_ReportStart()
'
    Dim j As Integer, wIndex As Integer
    Dim ws01 As String, wsS As String
    
    Dim w開始年月日 As Date, wdate As Date
    Dim wsNengetu As String
    Dim w推移表区分 As String
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
    Me.PageHeader.Controls("L_金融リストラ番号").Visible = True
    Me.PageHeader.Controls("H00_金融リストラ番号").Visible = True
'
    '----------------------------------------------------------------
    '                           ** 見出し **
    '----------------------------------------------------------------
    出力日 = Now
    企業名 = GCoName
'
    H00_金融リストラ番号 = GRpt.金融
'
    GRpt.テキスト_02 = GRpt.テキスト_01
    
    wdate = C年月日.平成To西暦("年月日", GRpt.テキスト_02)
    'GRpt.テキスト_01 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "e")
    '2019/01/15 日付入力区分 仕様変更
    If G基本情報.日付入力区分 = "0" Then
    '和暦
        If Len(GRpt.テキスト_01) <= 2 Then
            GRpt.テキスト_01 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "e")
        Else
            GRpt.テキスト_01 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "y")
        End If
    Else
    '西暦
            GRpt.テキスト_01 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "y")
    End If

    wsNengetu = GRpt.テキスト_01
    GRpt.推移 = "月次"
    
    L_帳票名.Caption = GRpt.帳票名 & " -" & GRpt.テキスト_02 & " -"
    
    If GRpt.千円単位 = 1 Then
        w分母 = 1000
        L_単位.Caption = "（千円単位）"
    Else
        w分母 = 1
        L_単位 = "（円単位）"
    End If
'
    '帳票指示
    wsS = ""
    If GRpt.借入金管理区分 = P8.FCDbl(XMXA020_区分("借入金管理区分", "決算用")) Then
        wsS = wsS & "帳票指示:決算用"
    Else
        wsS = wsS & "帳票指示:管理用"
    End If
    Me.PageHeader.Controls("L_帳票指示").Caption = wsS
'
    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    Call MBD020_借入金ワークテーブル作成("DBDA010_借入金")
    Call MRB010_標準入力借入残高表("DCIA010_借入金ワーク")       '07/02/18 V180
    Call MRB010_手入力借入残高表("DCIA010_借入金ワーク")         '07/02/09 V180
'
    'w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "平成")
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

    wIndex = 1
    For j = 1 To 12
        If GRpt.テキスト_02 = w推移表タイトル.X番目年月(j) Then
                wIndex = j
            Exit For
        End If
    Next
'
    ws01 = Right("00" & CStr(wIndex), 2)
    
    wstr = ""
    If GRpt.S_金融 = "分類1" Then
        wstr = wstr & "SELECT G.金融機関番号"
        wstr = wstr & " FROM (DCDA010_借入残高推移表結果 As Z"
        wstr = wstr & " INNER JOIN DBDA010_借入金 As K"
        wstr = wstr & " ON Z.借入番号 = K.借入番号)"
        wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
        wstr = wstr & " ON K.銀行番号 = G.銀行番号"
        wstr = wstr & " GROUP BY G.金融機関番号"
    Else
        wstr = wstr & "SELECT G.銀行番号"
        wstr = wstr & " FROM (DCDA010_借入残高推移表結果 As Z"
        wstr = wstr & " INNER JOIN DBDA010_借入金 As K"
        wstr = wstr & " ON Z.借入番号 = K.借入番号)"
        wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
        wstr = wstr & " ON K.銀行番号 = G.銀行番号"
        wstr = wstr & " GROUP BY G.銀行番号"
    End If
    wstr = wstr & " HAVING Sum(Z.融資_" & ws01 & ")<>0"
    wstr = wstr & " OR Sum(Z.元金_" & ws01 & ")<>0"
    wstr = wstr & " OR Sum(Z.利息_" & ws01 & ")<>0"
    wstr = wstr & " OR Sum(Z.返済_" & ws01 & ")<>0"
    wstr = wstr & " OR Sum(Z.残高_" & ws01 & ")<>0"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
        wiGCnt = wRs.RecordCount
    End If
    wRs.Close
    Set wRs = Nothing
'
    Call グラフデータ設定
'
    '集計 ワークテーブル作成
    Call 指定年月金融機関集計(wIndex)
'
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
'
    '----------------------------------------------------------------
    '          ** グループ ・ Where ・ OrderＢｙ **
    '----------------------------------------------------------------
    '----------------------------------------------------------------
    '                       ** レコード　ソース **
    '----------------------------------------------------------------
    '** レコード　ソース **
    wstr = "Select "
    wstr = wstr & "W.科目番号 As I_銀行番号,"
    wstr = wstr & "W.科目名 As I_銀行名,"
    wstr = wstr & "W.コード_001 As I_カウント,"
    wstr = wstr & "W.コード_002 As I_融資金額,"
    wstr = wstr & "W.コード_003 As I_融資,"
    wstr = wstr & "W.コード_004 As I_元金,"
    wstr = wstr & "W.コード_005 As I_利息,"
    wstr = wstr & "W.コード_006 As I_返済,"
    wstr = wstr & "W.コード_007 As I_残高"
    wstr = wstr & " FROM DCXA020_帳票作成ワーク As W"
    wstr = wstr & " Order by W.科目番号"
    Me.DataControl1.Source = wstr
    
    wiCnt = 1
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
' ActiveReport_ReportEnd
'------------------------------------------------
Private Sub ActiveReport_ReportEnd()
'
    'FBA010_帳票範囲指定.メッセージ = ""
    'FBA010_帳票範囲指定.メッセージ.Refresh
'
    ' =========================================
    '           　 ボタン制御
    ' =========================================
    'FBA010_帳票範囲指定.実行.Enabled = True
    'FBA010_帳票範囲指定.閉じる.Enabled = True
'
    'FBA010_帳票範囲指定.拡張.SetFocus
'
    ' =========================================
    '               LOG_WRITE
    ' =========================================
    GLogStr = "年月=" & GRpt.テキスト_02
    'Call MXA030_LOG_WRITE(GRpt.帳票名, "帳票", GLogStr)
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

'------------------------------------------------
' Detail_BeforePrint
'------------------------------------------------
Private Sub Detail_BeforePrint()
'
    Dim wi01 As Integer
'
    wdYusi(wiCnt) = P8.FCDbl(Me.Detail.Controls("I_融資金額"))
    Me.Detail.Controls("I_融資金額") = Format(wdYusi(wiCnt) / w分母, "#,##0")
    wdPYusi(wiCnt) = Format(P8.FFix(P8.FCDiv(wdYusi(wiCnt), wdYusi(0)) * 10000) / 100, "#,##0.00")
    Me.Detail.Controls("I_融資金額K") = wdPYusi(wiCnt)
    
    wdTokiYusi(wiCnt) = P8.FCDbl(Me.Detail.Controls("I_融資"))
    Me.Detail.Controls("I_融資") = Format(wdTokiYusi(wiCnt) / w分母, "#,##0")
    wdPTokiYusi(wiCnt) = Format(P8.FFix(P8.FCDiv(wdTokiYusi(wiCnt), wdTokiYusi(0)) * 10000) / 100, "#,##0.00")
    Me.Detail.Controls("I_融資K") = wdPTokiYusi(wiCnt)
    
    wdGankin(wiCnt) = P8.FCDbl(Me.Detail.Controls("I_元金"))
    Me.Detail.Controls("I_元金") = Format(wdGankin(wiCnt) / w分母, "#,##0")
    wdPGankin(wiCnt) = Format(P8.FFix(P8.FCDiv(wdGankin(wiCnt), wdGankin(0)) * 10000) / 100, "#,##0.00")
    Me.Detail.Controls("I_元金K") = wdPGankin(wiCnt)
    
    wdRisoku(wiCnt) = P8.FCDbl(Me.Detail.Controls("I_利息"))
    Me.Detail.Controls("I_利息") = Format(wdRisoku(wiCnt) / w分母, "#,##0")
    wdPRisoku(wiCnt) = Format(P8.FFix(P8.FCDiv(wdRisoku(wiCnt), wdRisoku(0)) * 10000) / 100, "#,##0.00")
    Me.Detail.Controls("I_利息K") = wdPRisoku(wiCnt)
    
    wdHensai(wiCnt) = P8.FCDbl(Me.Detail.Controls("I_返済"))
    Me.Detail.Controls("I_返済") = Format(wdHensai(wiCnt) / w分母, "#,##0")
    wdPHensai(wiCnt) = Format(P8.FFix(P8.FCDiv(wdHensai(wiCnt), wdHensai(0)) * 10000) / 100, "#,##0.00")
    Me.Detail.Controls("I_返済K") = wdPHensai(wiCnt)
    
    wdzan(wiCnt) = P8.FCDbl(Me.Detail.Controls("I_残高"))
    Me.Detail.Controls("I_残高") = Format(wdzan(wiCnt) / w分母, "#,##0")
    wdPZan(wiCnt) = Format(P8.FFix(P8.FCDiv(wdzan(wiCnt), wdzan(0)) * 10000) / 100, "#,##0.00")
    Me.Detail.Controls("I_残高K") = wdPZan(wiCnt)

    Me.Detail.Controls("I_返済率") = Format(Round(P8.FCDiv(wdYusi(wiCnt) - wdzan(wiCnt), wdYusi(wiCnt)) * 100, 3), "#,##0.00")
'
    wi01 = CInt(Right(CStr(wiCnt), 1))
    Me.Detail.Controls("Shape_L").BackColor = pColor(wi01)
'
End Sub

'------------------------------------------------
' Detail_AfterPrint
'------------------------------------------------
Private Sub Detail_AfterPrint()
    wiCnt = wiCnt + 1
End Sub

'------------------------------------------------
' ReportFooter_BeforePrint
'------------------------------------------------
Private Sub ReportFooter_BeforePrint()
'
    Dim j As Integer, wi01 As Integer
    Dim wd01 As Double
    Dim wstr As String
    Dim ws01 As String
    Dim Yusi As Double, Yusizan As Double
'
    Yusi = P8.FCDbl(Me.ReportFooter.Controls("G90_融資金額"))
    Yusizan = P8.FCDbl(Me.ReportFooter.Controls("G90_残高"))
    
    Me.ReportFooter.Controls("G90_融資金額") = Format(P8.FCDbl(Me.ReportFooter.Controls("G90_融資金額")) / w分母, "#,##0")
    Me.ReportFooter.Controls("G90_融資") = Format(P8.FCDbl(Me.ReportFooter.Controls("G90_融資")) / w分母, "#,##0")
    Me.ReportFooter.Controls("G90_元金") = Format(P8.FCDbl(Me.ReportFooter.Controls("G90_元金")) / w分母, "#,##0")
    Me.ReportFooter.Controls("G90_利息") = Format(P8.FCDbl(Me.ReportFooter.Controls("G90_利息")) / w分母, "#,##0")
    Me.ReportFooter.Controls("G90_返済") = Format(P8.FCDbl(Me.ReportFooter.Controls("G90_返済")) / w分母, "#,##0")
    Me.ReportFooter.Controls("G90_残高") = Format(P8.FCDbl(Me.ReportFooter.Controls("G90_残高")) / w分母, "#,##0")

    Me.ReportFooter.Controls("G90_返済率") = Format(Round(P8.FCDiv(Yusi - Yusizan, Yusi) * 100, 3), "#,##0.00")
'
    Call グラフ作成
'
End Sub

'------------------------------------------------
' グラフデータ設定
'------------------------------------------------
Private Sub グラフデータ設定()
'
    On Error GoTo グラフデータ設定_ERR
'
    'グラフ
    ReDim ws_GrpName(wiGCnt)
    
    ReDim wdYusi(wiGCnt)
    ReDim wdTokiYusi(wiGCnt)
    ReDim wdGankin(wiGCnt)
    ReDim wdRisoku(wiGCnt)
    ReDim wdHensai(wiGCnt)
    ReDim wdzan(wiGCnt)
    
    ReDim wdPYusi(wiGCnt)
    ReDim wdPTokiYusi(wiGCnt)
    ReDim wdPGankin(wiGCnt)
    ReDim wdPRisoku(wiGCnt)
    ReDim wdPHensai(wiGCnt)
    ReDim wdPZan(wiGCnt)
'
    'カラー
    pColor(1) = RGB(255, 99, 71)    '赤
    pColor(2) = RGB(135, 206, 250)  '青
    pColor(3) = RGB(0, 255, 102)    '黄緑
    pColor(4) = RGB(255, 255, 0)    '黄色
    pColor(5) = RGB(224, 255, 255)  '水色
    pColor(6) = RGB(255, 160, 122)  'オレンジ
    pColor(7) = RGB(204, 204, 204)  '灰色
    pColor(8) = RGB(255, 0, 255)    'ピンク
    pColor(9) = RGB(154, 205, 50)   '緑
    pColor(0) = RGB(221, 160, 221)  '紫
'
    'グラフ
    cht1.ColumnCount = wiGCnt
    cht1.RowLabel = "融資金額"
    
    'cht2.ColumnCount = wiGCnt
    'cht2.RowLabel = "当期融資金額"
    
    cht3.ColumnCount = wiGCnt
    cht3.RowLabel = "元金額"
    
    cht4.ColumnCount = wiGCnt
    cht4.RowLabel = "利息額"
    
    cht5.ColumnCount = wiGCnt
    cht5.RowLabel = "返済額"
    
    cht6.ColumnCount = wiGCnt
    cht6.RowLabel = "融資残高"
'
    Exit Sub
'
'----------< ERROR ROUTINE >--------------------------------------------------------
グラフデータ設定_ERR:
    pERR_MES = pPROGRAM_ID + "/ グラフデータ設定() でエラー" + vbCrLf + vbCrLf + _
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
' 指定年月金融機関集計
'------------------------------------------------
Private Sub 指定年月金融機関集計(pIndex As Integer)
'
    Dim wRs2 As ADODB.Recordset
    Dim wstr2 As String
    
    Dim wd01 As Double
    Dim ws01 As String
'
    On Error GoTo 指定年月金融機関集計_ERR
'
    wiCnt = 1
    
    wstr = ""
    wstr = wstr & "SELECT 科目番号,科目名,コード_001,コード_002,コード_003,コード_004,コード_005,コード_006,コード_007"
    wstr = wstr & " FROM DCXA020_帳票作成ワーク"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        ws01 = Right("00" & CStr(pIndex), 2)
        
        wstr2 = ""
        If GRpt.S_金融 = "分類1" Then
            wstr2 = wstr2 & "SELECT G.金融機関番号,Count(G.金融機関番号) AS カウント,G.金融機関名,"
            wstr2 = wstr2 & "Sum(K.融資金額) AS 融資金額合計,"
            wstr2 = wstr2 & "Sum(Z.融資_" & ws01 & ") AS 融資_合計,"
            wstr2 = wstr2 & "Sum(Z.残高_" & ws01 & ") AS 残高_合計,"
            wstr2 = wstr2 & "Sum(Z.元金_" & ws01 & ") AS 元金_合計,"
            wstr2 = wstr2 & "Sum(Z.利息_" & ws01 & ") AS 利息_合計,"
            wstr2 = wstr2 & "Sum(Z.返済_" & ws01 & ") AS 返済_合計"
            wstr2 = wstr2 & " FROM (DCDA010_借入残高推移表結果 As Z"
            wstr2 = wstr2 & " INNER JOIN DBDA010_借入金 As K"
            wstr2 = wstr2 & " ON Z.借入番号 = K.借入番号)"
            wstr2 = wstr2 & " INNER JOIN DAAA040_銀行マスタ As G"
            wstr2 = wstr2 & " ON K.銀行番号 = G.銀行番号"
            wstr2 = wstr2 & " GROUP BY G.金融機関番号,G.金融機関名"
            wstr2 = wstr2 & " HAVING Sum(Z.融資_" & ws01 & ")<>0"
            wstr2 = wstr2 & " OR Sum(Z.元金_" & ws01 & ")<>0"
            wstr2 = wstr2 & " OR Sum(Z.利息_" & ws01 & ")<>0"
            wstr2 = wstr2 & " OR Sum(Z.返済_" & ws01 & ")<>0"
            wstr2 = wstr2 & " OR Sum(Z.残高_" & ws01 & ")<>0"
            wstr2 = wstr2 & " ORDER BY G.金融機関番号"
        Else
            wstr2 = wstr2 & "SELECT G.銀行番号,Count(G.銀行番号) AS カウント,G.銀行名,"
            wstr2 = wstr2 & "Sum(K.融資金額) AS 融資金額合計,"
            wstr2 = wstr2 & "Sum(Z.融資_" & ws01 & ") AS 融資_合計,"
            wstr2 = wstr2 & "Sum(Z.残高_" & ws01 & ") AS 残高_合計,"
            wstr2 = wstr2 & "Sum(Z.元金_" & ws01 & ") AS 元金_合計,"
            wstr2 = wstr2 & "Sum(Z.利息_" & ws01 & ") AS 利息_合計,"
            wstr2 = wstr2 & "Sum(Z.返済_" & ws01 & ") AS 返済_合計"
            wstr2 = wstr2 & " FROM (DCDA010_借入残高推移表結果 As Z"
            wstr2 = wstr2 & " INNER JOIN DBDA010_借入金 As K"
            wstr2 = wstr2 & " ON Z.借入番号 = K.借入番号)"
            wstr2 = wstr2 & " INNER JOIN DAAA040_銀行マスタ As G"
            wstr2 = wstr2 & " ON K.銀行番号 = G.銀行番号"
            wstr2 = wstr2 & " GROUP BY G.銀行番号,G.銀行名"
            wstr2 = wstr2 & " HAVING Sum(Z.融資_" & ws01 & ")<>0"
            wstr2 = wstr2 & " OR Sum(Z.元金_" & ws01 & ")<>0"
            wstr2 = wstr2 & " OR Sum(Z.利息_" & ws01 & ")<>0"
            wstr2 = wstr2 & " OR Sum(Z.返済_" & ws01 & ")<>0"
            wstr2 = wstr2 & " OR Sum(Z.残高_" & ws01 & ")<>0"
            wstr2 = wstr2 & " ORDER BY G.銀行番号"
        End If
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
        Do Until wRs2.eof
            wRs.AddNew
                
                If GRpt.S_金融 = "分類1" Then
                    ws_GrpName(wiCnt) = wRs2("金融機関名")
                    wRs("科目番号") = wRs2("金融機関番号")
                    wRs("科目名") = wRs2("金融機関名")
                Else
                    ws_GrpName(wiCnt) = wRs2("銀行名")
                    wRs("科目番号") = wRs2("銀行番号")
                    wRs("科目名") = wRs2("銀行名")
                End If
                
                wRs("コード_001") = wRs2("カウント")
                
                wd01 = P8.FCDbl(wRs2("融資金額合計"))
                wRs("コード_002") = wd01
                wdYusi(0) = wdYusi(0) + wd01
                
                wd01 = P8.FCDbl(wRs2("融資_合計"))
                wRs("コード_003") = wd01
                wdTokiYusi(0) = wdTokiYusi(0) + wd01
                
                wd01 = P8.FCDbl(wRs2("元金_合計"))
                wRs("コード_004") = wd01
                wdGankin(0) = wdGankin(0) + wd01
                
                wd01 = P8.FCDbl(wRs2("利息_合計"))
                wRs("コード_005") = wd01
                wdRisoku(0) = wdRisoku(0) + wd01
                
                wd01 = P8.FCDbl(wRs2("返済_合計"))
                wRs("コード_006") = wd01
                wdHensai(0) = wdHensai(0) + wd01
                
                wd01 = P8.FCDbl(wRs2("残高_合計"))
                wRs("コード_007") = wd01
                wdzan(0) = wdzan(0) + wd01
                
                wiCnt = wiCnt + 1
            
            wRs.Update
            
            wRs2.MoveNext
        Loop
        wRs2.Close
        Set wRs2 = Nothing

    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >--------------------------------------------------------
指定年月金融機関集計_ERR:
    pERR_MES = pPROGRAM_ID + "/ 指定年月金融機関集計() でエラー" + vbCrLf + vbCrLf + _
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
' グラフ作成
'------------------------------------------------
Private Sub グラフ作成()
'
    Dim j As Integer, wi01 As Integer
    Dim wdMaxdata(4) As Double
    Dim wlTani(4) As Long
    Dim wd01 As Double
    Dim wstr As String
    Dim ws01 As String
    Dim Yusi As Double, Yusizan As Double
'
    On Error GoTo グラフ作成_ERR
'
    'y軸 単位
    For j = 0 To 4
        wdMaxdata(j) = 0
        wlTani(j) = 1
    Next j

    cht1.RowLabel = "融資金額　（円単位）"
    cht3.RowLabel = "元金額　（円単位）"
    cht4.RowLabel = "利息額　（円単位）"
    cht5.RowLabel = "返済額　（円単位）"
    cht6.RowLabel = "融資残高　（円単位）"
    
    For j = 1 To wiGCnt
        If wdMaxdata(0) < wdYusi(j) Then
            wdMaxdata(0) = wdYusi(j)
        End If
        If wdMaxdata(1) < wdGankin(j) Then
            wdMaxdata(1) = wdGankin(j)
        End If
        If wdMaxdata(2) < wdRisoku(j) Then
            wdMaxdata(2) = wdRisoku(j)
        End If
        If wdMaxdata(3) < wdHensai(j) Then
            wdMaxdata(3) = wdHensai(j)
        End If
        If wdMaxdata(4) < wdzan(j) Then
            wdMaxdata(4) = wdzan(j)
        End If
    Next j

    '6桁以上で桁が溢れるため
    If wdMaxdata(0) / 1000000 >= 1000000 Then  '1,000,000,000,000
        wlTani(0) = 1000000000
        cht1.RowLabel = "融資金額　（10億単位）"
    ElseIf wdMaxdata(0) / 1000 >= 1000000 Then '1,000,000,000
        wlTani(0) = 1000000
        cht1.RowLabel = "融資金額　（百万単位）"
    ElseIf wdMaxdata(0) >= 1000000 Then         '1,000,000
        wlTani(0) = 1000
        cht1.RowLabel = "融資金額　（千円単位）"
    Else
        If GRpt.千円単位 = 1 Then
            wlTani(0) = 1000
            cht1.RowLabel = "融資金額　（千円単位）"
        End If
    End If
    
    If wdMaxdata(1) / 1000000 >= 1000000 Then
        wlTani(1) = 1000000000
        cht3.RowLabel = "元金額　（10億単位）"
    ElseIf wdMaxdata(1) / 1000 >= 1000000 Then
        wlTani(1) = 1000000
        cht3.RowLabel = "元金額　（百万単位）"
    ElseIf wdMaxdata(1) >= 1000000 Then
        wlTani(1) = 1000
        cht3.RowLabel = "元金額　（千円単位）"
    Else
        If GRpt.千円単位 = 1 Then
            wlTani(1) = 1000
            cht3.RowLabel = "元金額　（千円単位）"
        End If
    End If
    
    If wdMaxdata(2) / 1000000 >= 1000000 Then
        wlTani(2) = 1000000000
        cht4.RowLabel = "利息額　（10億単位）"
    ElseIf wdMaxdata(2) / 1000 >= 1000000 Then
        wlTani(2) = 1000000
        cht4.RowLabel = "利息額　（百万単位）"
    ElseIf wdMaxdata(2) >= 1000000 Then
        wlTani(2) = 1000
        cht4.RowLabel = "利息額　（千円単位）"
    Else
        If GRpt.千円単位 = 1 Then
            wlTani(2) = 1000
            cht4.RowLabel = "利息額　（千円単位）"
        End If
    End If
    
    If wdMaxdata(3) / 1000000 >= 1000000 Then
        wlTani(3) = 1000000000
        cht5.RowLabel = "返済額　（10億単位）"
    ElseIf wdMaxdata(3) / 1000 >= 1000000 Then
        wlTani(3) = 1000000
        cht5.RowLabel = "返済額　（百万単位）"
    ElseIf wdMaxdata(3) >= 1000000 Then
        wlTani(3) = 1000
        cht5.RowLabel = "返済額　（千円単位）"
    Else
        If GRpt.千円単位 = 1 Then
            wlTani(3) = 1000
            cht5.RowLabel = "返済額　（千円単位）"
        End If
    End If
    
    If wdMaxdata(4) / 1000000 >= 1000000 Then
        wlTani(4) = 1000000000
        cht6.RowLabel = "融資残高　（10億単位）"
    ElseIf wdMaxdata(4) / 1000 >= 1000000 Then
        wlTani(4) = 1000000
        cht6.RowLabel = "融資残高　（百万単位）"
    ElseIf wdMaxdata(4) >= 1000000 Then
        wlTani(4) = 1000
        cht6.RowLabel = "融資残高　（千円単位）"
    Else
        If GRpt.千円単位 = 1 Then
            wlTani(4) = 1000
            cht6.RowLabel = "融資残高　（千円単位）"
        End If
    End If
'
'   '目盛り線の設定
'   With cht1.Plot.Axis(1).ValueScale
'      wd01 = wdMaxdata(0) / wlTani(0)
'      If Right(wd01, 1) <> 0 Then
'        wd01 = P8.FFix((wd01 + 1000) / 1000) * 1000
'      End If
'
'      .Auto = False                          '自動設定を解除
'      .Maximum = wd01                        '最大値(調整)
'      .Minimum = 0                           '最小値
'      .MajorDivision = 5                    'メモリ数
'   End With
'
'   With cht3.Plot.Axis(1).ValueScale
'      wd01 = wdMaxdata(1) / wlTani(1)
'      If Right(wd01, 1) <> 0 Then
'        wd01 = P8.FFix((wd01 + 1000) / 1000) * 1000
'      End If
'
'      .Auto = False
'      .Maximum = wd01
'      .Minimum = 0
'      .MajorDivision = 5
'   End With
'
'   With cht4.Plot.Axis(1).ValueScale
'      wd01 = wdMaxdata(2) / wlTani(2)
'      If Right(wd01, 1) <> 0 Then
'        wd01 = P8.FFix((wd01 + 100) / 100) * 100
'      End If
'
'      .Auto = False
'      .Maximum = wd01
'      .Minimum = 0
'      .MajorDivision = 5
'   End With
'
'   With cht5.Plot.Axis(1).ValueScale
'      wd01 = wdMaxdata(3) / wlTani(3)
'      If Right(wd01, 1) <> 0 Then
'        wd01 = P8.FFix((wd01 + 100) / 100) * 100
'      End If
'
'      .Auto = False
'      .Maximum = wd01
'      .Minimum = 0
'      .MajorDivision = 5
'   End With
'
'   With cht6.Plot.Axis(1).ValueScale
'      wd01 = wdMaxdata(4) / wlTani(4)
'      If Right(wd01, 1) <> 0 Then
'        wd01 = P8.FFix((wd01 + 100) / 100) * 100
'      End If
'
'      .Auto = False
'      .Maximum = wd01
'      .Minimum = 0
'      .MajorDivision = 5
'   End With
'
    For j = 1 To wiGCnt
        cht1.Column = j
        'cht1.ColumnLabel = ws_GrpName(j) & wdYusi(j) / wlTani(0)
        cht1.Data = P8.FRound(wdYusi(j) / wlTani(0), 1)
        
        'cht2.Column = j
        'cht2.ColumnLabel = ws_GrpName(j) & wdPTokiYusi(j)
        'cht2.Data = wdTokiYusi(j)
        
        cht3.Column = j
        'cht3.ColumnLabel = ws_GrpName(j) & wdGankin(j) / wlTani(1)
        cht3.Data = P8.FRound(wdGankin(j) / wlTani(1), 1)
        
        cht4.Column = j
        'cht4.ColumnLabel = ws_GrpName(j) & wdRisoku(j) / wlTani(2)
        cht4.Data = P8.FRound(wdRisoku(j) / wlTani(2), 1)
        
        cht5.Column = j
        'cht5.ColumnLabel = ws_GrpName(j) & wdHensai(j) / wlTani(3)
        cht5.Data = P8.FRound(wdHensai(j) / wlTani(3), 1)
        
        cht6.Column = j
        'cht6.ColumnLabel = ws_GrpName(j) & wdZan(j) / wlTani(4)
        cht6.Data = P8.FRound(wdzan(j) / wlTani(4), 1)
        
        
        wi01 = CInt(Right(CStr(j), 1))
        Select Case wi01
        Case 1
            cht1.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 99, 71     '赤
            'cht2.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 99, 71
            cht3.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 99, 71
            cht4.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 99, 71
            cht5.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 99, 71
            cht6.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 99, 71
        Case 2
            cht1.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 135, 206, 250   '青
            'cht2.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 135, 206, 250
            cht3.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 135, 206, 250
            cht4.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 135, 206, 250
            cht5.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 135, 206, 250
            cht6.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 135, 206, 250
        Case 3
            cht1.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 0, 255, 102    '黄緑
            'cht2.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 0, 255, 102
            cht3.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 0, 255, 102
            cht4.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 0, 255, 102
            cht5.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 0, 255, 102
            cht6.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 0, 255, 102
        Case 4
            cht1.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 255, 0   '黄色
            'cht2.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 255, 0
            cht3.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 255, 0
            cht4.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 255, 0
            cht5.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 255, 0
            cht6.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 255, 0
        Case 5
            cht1.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 224, 255, 255  '水色
            'cht2.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 224, 255, 255
            cht3.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 224, 255, 255
            cht4.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 224, 255, 255
            cht5.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 224, 255, 255
            cht6.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 224, 255, 255
        Case 6
            cht1.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 160, 122   'オレンジ
            'cht2.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 160, 122
            cht3.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 160, 122
            cht4.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 160, 122
            cht5.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 160, 122
            cht6.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 160, 122
        Case 7
            cht1.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 204, 204, 204    '灰色
            'cht2.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 204, 204, 204
            cht3.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 204, 204, 204
            cht4.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 204, 204, 204
            cht5.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 204, 204, 204
            cht6.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 204, 204, 204
        Case 8
            cht1.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 0, 255 'ピンク
            'cht2.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 0, 255
            cht3.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 0, 255
            cht4.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 0, 255
            cht5.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 0, 255
            cht6.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 255, 0, 255
        Case 9
            cht1.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 154, 205, 50 '緑
            'cht2.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 154, 205, 50
            cht3.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 154, 205, 50
            cht4.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 154, 205, 50
            cht5.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 154, 205, 50
            cht6.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 154, 205, 50
        Case 0
            cht1.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 221, 160, 221  '紫
            'cht2.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 221, 160, 221
            cht3.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 221, 160, 221
            cht4.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 221, 160, 221
            cht5.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 221, 160, 221
            cht6.Plot.SeriesCollection(j).DataPoints(-1).Brush.FillColor.Set 221, 160, 221
        End Select
    Next j
'
    Exit Sub
'
'----------< ERROR ROUTINE >--------------------------------------------------------
グラフ作成_ERR:
    pERR_MES = pPROGRAM_ID + "/ グラフ作成() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub
