VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RLA010_ログ一覧表 
   Caption         =   "ログ一覧表"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   Icon            =   "RLA010_ログ一覧表.dsx":0000
   StartUpPosition =   3  'Windows の既定値
   _ExtentX        =   26882
   _ExtentY        =   7303
   SectionData     =   "RLA010_ログ一覧表.dsx":0ECA
End
Attribute VB_Name = "RLA010_ログ一覧表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "RLA010_ログ一覧表"

'------------------------------------------------
' ActiveReport_ReportStart
'------------------------------------------------
Private Sub ActiveReport_ReportStart()
'
    Dim wstr As String
    Dim ws01 As String, ws02 As String
'
    On Error GoTo ActiveReport_ReportStart_ERR
'
    '----------------------------------------------------------------
    '                         ** 初期設定 **
    '----------------------------------------------------------------
    'Connection
    '---------------------------
    Me.DataControl1.Connection = GDb
'
    '用紙セット
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORPortrait
    
    '----------------------------------------------------------------
    '                           ** 見出し **
    '----------------------------------------------------------------
    出力日 = Now
    
    '----------------------------------------------------------------
    '                       ** レコード　ソース **
    '----------------------------------------------------------------
    GWhere = ""
    GWhere = " Where (1=1)"
    If GRpt.コンボ_01 <> "" Then
        GWhere = GWhere + " And マイコンピュータ = '" & GRpt.コンボ_01 & "'"
    End If
    
    ws01 = C年月日.平成To西暦("年月日", GRpt.コンボ_02)
    If ws01 <> "0" Then
        GWhere = GWhere + " And Format(更新日付,'yyyy/mm/dd') >= '" & Format(ws01, "yyyy/mm/dd") & "'"
    End If
    
    ws02 = C年月日.平成To西暦("年月日", GRpt.コンボ_03)
    If ws02 <> "0" Then
        GWhere = GWhere + " And Format(更新日付,'yyyy/mm/dd') <= '" & Format(ws02, "yyyy/mm/dd") & "'"
    End If
'
    wstr = ""
    wstr = wstr + "Select"
    wstr = wstr + " Format(更新日付,'ee年mm月dd日 hh:nn:ss') As I_更新日付,"
    wstr = wstr + " マイコンピュータ As I_マイコンピュータ,"
    wstr = wstr + " IIf(ログ区分 = '0','端末の開始',"
    wstr = wstr + "  IIf(ログ区分 = '1','- 企業名の登録',"
    wstr = wstr + "   IIf(ログ区分 = '2','- 企業処理の開始',"
    wstr = wstr + "    IIf(ログ区分 = '3','- 企業処理の終了',"
    wstr = wstr + "     IIf(ログ区分 = '4','端末の終了',"
    wstr = wstr + " IIf(ログ区分 = '5','  - 会議計画の開始',"
    wstr = wstr + "  IIf(ログ区分 = '6','  - 会議計画の終了',''))))))) As I_ログ区分,"
    wstr = wstr + " IIf(処理内容 = '1','新規追加',"
    wstr = wstr + "  IIf(処理内容 = '2','修正',"
    wstr = wstr + "   IIf(処理内容 = '3','削除',"
    wstr = wstr + "    IIf(処理内容 = '4','復元',"
    wstr = wstr + "     IIf(処理内容 = '5','バックアップ',"
    wstr = wstr + "      IIf(処理内容 = '6','フラグ解除','')))))) As I_処理内容,"
    wstr = wstr + " 企業名Key As I_企業名Key,"
    wstr = wstr + " 企業名 As I_企業名"
    wstr = wstr + " From DCLA010_ログファイル"
    
    GOrder = " Order BY 更新日付,マイコンピュータ"
    wstr = wstr + GWhere + GOrder
    
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
    'FBC010_ログ一覧表.メッセージ = ""
    'FBC010_ログ一覧表.メッセージ.Refresh
'
    ' =========================================
    '           　 ボタン制御
    ' =========================================
    'FBC010_ログ一覧表.実行.Enabled = True
    'FBC010_ログ一覧表.閉じる.Enabled = True
'
    'FBC010_ログ一覧表.拡張.SetFocus
'
End Sub

'------------------------------------------------
' ActiveReport_NoData
'------------------------------------------------
Private Sub ActiveReport_NoData()
'
    'FBC010_ログ一覧表.メッセージ = "出力すべきデータはありません"
    'FBC010_ログ一覧表.メッセージ.Refresh
    '
    Me.Cancel
    DoEvents
'
    ' =========================================
    '           　 ボタン制御
    ' =========================================
    'FBC010_ログ一覧表.実行.Enabled = True
    'FBC010_ログ一覧表.閉じる.Enabled = True
'
    'FBC010_ログ一覧表.拡張.SetFocus
'
    Unload Me
'
End Sub

'------------------------------------------------
' ActiveReport_Error
'------------------------------------------------
Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
'
    'FBC010_ログ一覧表.メッセージ = "出力できませんでした"
    'FBC010_ログ一覧表.メッセージ.Refresh
    
    Me.Cancel
    DoEvents

    ' =========================================
    '           　 ボタン制御
    ' =========================================
    'FBC010_ログ一覧表.実行.Enabled = True
    'FBC010_ログ一覧表.閉じる.Enabled = True
'
    Unload Me
'
End Sub


