VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RDG020_残高表グラフ 
   Caption         =   "金融機関別残高表"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12690
   Icon            =   "RDG020_残高表グラフ.dsx":0000
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   _ExtentX        =   22384
   _ExtentY        =   16510
   SectionData     =   "RDG020_残高表グラフ.dsx":0ECA
End
Attribute VB_Name = "RDG020_残高表グラフ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "RDG020_残高表グラフ"
'
Dim wRs As ADODB.Recordset

Dim wstr As String
Dim wWhere As String

Dim w分母 As Integer

'グループ集計
Dim wdGZan As Double, wdGTanki As Double, wdGTyoki As Double, wdGShasai As Double
Dim wdPZan As Double, wdPTanki As Double, wdPTyoki As Double, wdPShasai As Double

Dim w推移表タイトル As MAA910_推移表タイトル
'
'------------------------------------------------
' ActiveReport_ReportStart
'------------------------------------------------
Private Sub ActiveReport_ReportStart()
'
    Dim wsS As String
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
    '          ** グループ ・ Where ・ OrderＢｙ **
    '----------------------------------------------------------------
    
    '----------------------------------------------------------------
    '                       ** レコード　ソース **
    '----------------------------------------------------------------
    '** レコード　ソース **
    wstr = "Select "
    wstr = wstr & " Sum(W.コード_007) AS 残高合計,"
    wstr = wstr & " Sum(W.コード_008) AS 短期合計,"
    wstr = wstr & " Sum(W.コード_009) AS 長期合計,"
    wstr = wstr & " Sum(W.コード_010) AS 社債合計"
    wstr = wstr & " FROM DCXA020_帳票作成ワーク AS W"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
        wdGZan = P8.FCDbl(wRs("残高合計"))
        wdGTanki = P8.FCDbl(wRs("短期合計"))
        wdGTyoki = P8.FCDbl(wRs("長期合計"))
        wdGShasai = P8.FCDbl(wRs("社債合計"))
    End If
    wRs.Close
    Set wRs = Nothing
'
    wstr = "Select "
    wstr = wstr & "W.科目番号 As I_銀行番号,"
    wstr = wstr & "W.科目名 As I_銀行名,"
    wstr = wstr & "W.コード_001 As I_カウント,"
    wstr = wstr & "W.コード_002 As I_融資金額,"
    wstr = wstr & "W.コード_003 As I_融資,"
    wstr = wstr & "W.コード_004 As I_元金,"
    wstr = wstr & "W.コード_005 As I_利息,"
    wstr = wstr & "W.コード_006 As I_返済,"
    wstr = wstr & "W.コード_007 As I_残高,"
    wstr = wstr & "W.コード_008 As I_短期,"
    wstr = wstr & "W.コード_009 As I_長期,"
    wstr = wstr & "W.コード_010 As I_社債,"
    wstr = wstr & "W.コード_011 As I_短期C,"
    wstr = wstr & "W.コード_012 As I_長期C,"
    wstr = wstr & "W.コード_013 As I_社債C"
    wstr = wstr & " FROM DCXA020_帳票作成ワーク As W"
    wstr = wstr & " Order by W.科目番号"
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
    'CSVファイル パラメータ:GInt1
    If GRpt.CSV = 1 Then
        Call MX040_CsvOut_KARISUII(w推移表タイトル)
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
    Dim wd01 As Double
'
    wd01 = P8.FCDbl(Me.Detail.Controls("I_残高"))
    Me.Detail.Controls("I_残高") = Format(wd01 / w分母, "#,##0")
    wdPZan = P8.FFix(P8.FCDiv(wd01, wdGZan) * 10000) / 100
    Me.Detail.Controls("I_残高K") = Format(wdPZan, "#,##0.00")

    wd01 = P8.FCDbl(Me.Detail.Controls("I_短期"))
    Me.Detail.Controls("I_短期") = Format(wd01 / w分母, "#,##0")
    wdPTanki = P8.FFix(P8.FCDiv(wd01, wdGTanki) * 10000) / 100
    Me.Detail.Controls("I_短期K") = Format(wdPTanki, "#,##0.00")
    
    wd01 = P8.FCDbl(Me.Detail.Controls("I_長期"))
    Me.Detail.Controls("I_長期") = Format(wd01 / w分母, "#,##0")
    wdPTyoki = P8.FFix(P8.FCDiv(wd01, wdGTyoki) * 10000) / 100
    Me.Detail.Controls("I_長期K") = Format(wdPTyoki, "#,##0.00")
    
    wd01 = P8.FCDbl(Me.Detail.Controls("I_社債"))
    Me.Detail.Controls("I_社債") = Format(wd01 / w分母, "#,##0")
    wdPShasai = P8.FFix(P8.FCDiv(wd01, wdGShasai) * 10000) / 100
    Me.Detail.Controls("I_社債K") = Format(wdPShasai, "#,##0.00")
'
End Sub

'------------------------------------------------
' ReportFooter_BeforePrint
'------------------------------------------------
Private Sub ReportFooter_BeforePrint()
'
    Me.ReportFooter.Controls("G90_残高") = Format(P8.FCDbl(Me.ReportFooter.Controls("G90_残高")) / w分母, "#,##0")
    Me.ReportFooter.Controls("G90_短期") = Format(P8.FCDbl(Me.ReportFooter.Controls("G90_短期")) / w分母, "#,##0")
    Me.ReportFooter.Controls("G90_長期") = Format(P8.FCDbl(Me.ReportFooter.Controls("G90_長期")) / w分母, "#,##0")
    Me.ReportFooter.Controls("G90_社債") = Format(P8.FCDbl(Me.ReportFooter.Controls("G90_社債")) / w分母, "#,##0")
'
End Sub
