VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RDA100_借入金台帳 
   Caption         =   "借入金台帳"
   ClientHeight    =   8460
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12195
   Icon            =   "RDA100_借入金台帳.dsx":0000
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   WindowState     =   2  '最大化
   _ExtentX        =   21511
   _ExtentY        =   14923
   SectionData     =   "RDA100_借入金台帳.dsx":0ECA
End
Attribute VB_Name = "RDA100_借入金台帳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "RDA010_借入明細表"

Dim wstr As String, w円単位 As String, wsTbl As String, wsTbl2 As String
Dim w分母 As Integer
Dim FLG_TR As Boolean
'
'------------------------------------------------
' ActiveReport_ReportStart
'------------------------------------------------
Private Sub ActiveReport_ReportStart()
'
    Dim wRs As ADODB.Recordset
    Dim wWhere As String

    Dim w借入データ As MAA910_借入金
    Dim wdHRiritu As Double
    Dim w金融リストラ As String

    Dim wdKinri As Double
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
'
'
    '借入明細表 or 貸付明細表
    Select Case GRpt.帳票名
    Case "借入金台帳"
        wsTbl = "DBDA010_借入金"
        wsTbl2 = "DBDA010_借入金明細TR"

        Me.PageHeader.Controls("L_帳票名").Caption = "借入金台帳"

    Case "貸付明細表"
        wsTbl = "DBDA010_貸付金"
        wsTbl2 = "DBDA010_貸付金明細TR"

        Me.PageHeader.Controls("L_帳票名").Caption = "貸付金台帳"

    End Select
'
    If G金利SM = True Then
        L_帳票名.Caption = " " & GRpt.帳票名 & " - 金利SM - "
    End If
'
    '通常 or 手入力 の書式設定は↓
'
    '----------------------------------------------------------------
    '                           ** 見出し **
    '----------------------------------------------------------------
    出力日 = Now
    企業名 = GCoName

    I_借入番号 = GRpt.コンボ_01

    '----------------------------------------------------------------
    '          ** グループ ・ Where ・ OrderＢｙ **
    '----------------------------------------------------------------
    wWhere = ""

    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    wstr = ""
    wstr = wstr + "Select K.借入番号 AS I_借入番号,"
    wstr = wstr + "K.借入内容 AS I_借入内容,"
    wstr = wstr + "K.手入力区分 AS T_手入力区分,"
    wstr = wstr + " IIF(K.sm区分=0,'OFF','ON') As I_SM区分,"
    wstr = wstr + "K.金融リストラ番号 AS I_金融リストラ番号,"
    wstr = wstr + "K.銀行番号 AS I_銀行番号,"
    wstr = wstr + "IIF(K.支払日=31,'月末',(FORMAT(K.支払日,'##日'))) AS I_支払日,"
    
    wstr = wstr + " IIF(K.手入力区分 = " & P8.FCDbl(XMXA020_区分("日割計算区分", "自動計算")) & ",'自動計算','入力登録') As I_日割計算,"
    wstr = wstr + " IIF(K.手入力区分 = " & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",'標準登録','入力登録') As I_登録方法,"
    wstr = wstr + " IIF(K.営業日区分 = " & P8.FCDbl(XMXA020_区分("営業日", "翌営業日")) & ",'翌営業日','前営業日') As I_営業日,"
    wstr = wstr + " IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As I_利息区分,"
    wstr = wstr + " IIF(K.利息計算日数区分 = " & P8.FCDbl(XMXA020_区分("利息計算日数", "営業日数")) & ",'営業日数','固定日数') As I_利息計算日数,"
    wstr = wstr + " IIF(K.利息支払方法 = " & P8.FCDbl(XMXA020_区分("利息支払", "毎月")) & ",'毎月','一括') As I_利息支払方法,"
    wstr = wstr + " IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "控除無し")) & ",'控除無し',"
    wstr = wstr + " IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "実行日控除")) & ",'実行日控除',"
    wstr = wstr + " IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "最終返済日控除")) & ",'最終返済日控除',"
    wstr = wstr + " IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) & ",'実行日及び最終返済日控除','中間利払最終日控除')))) As I_利息控除区分,"
    wstr = wstr + " IIF(K.金利計算年間日数 = " & P8.FCDbl(XMXA020_区分("金利計算", "365日")) & ",'365日','360日') As I_金利計算年間日数,"
    wstr = wstr + " IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As I_金利種別,"
    wstr = wstr + " IIF(K.有担保フラグ=0,'無担保','有担保') As I_担保区分,"
    wstr = wstr & " IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期借入金','長期借入金') AS I_長短区分,"
    wstr = wstr & "IIF(K.設備フラグ=0,'運転資金','設備資金') AS I_設備フラグ,"
    wstr = wstr + "FORMAT(K.融資金額,'#,###,###,###,##0円') AS I_融資金額,"
    wstr = wstr + "FORMAT(K.利率,'0.00000') AS I_利率,"
    wstr = wstr + "FORMAT(K.実行日,'" & Gfmt年月日 & "') AS I_実行日,"
    wstr = wstr + "FORMAT(K.初回返済年月,'" & Gfmt年月 & "') AS I_初回返済年月,"
    wstr = wstr + "FORMAT(K.最終返済年月,'" & Gfmt年月 & "') AS I_最終返済年月,"
    wstr = wstr + "FORMAT(K.初回返済実行日,'" & Gfmt年月日 & "') AS I_初回返済実行日,"
    wstr = wstr + "FORMAT(K.最終返済実行日,'" & Gfmt年月日 & "') AS I_最終返済実行日,"
    
    wstr = wstr & " IIF(K.手入力区分 = " & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",FORMAT(K.金利初回年月,'" & Gfmt年月 & "'),'**') As I_金利初回年月,"
    
'    wstr = wstr + "FORMAT(K.金利初回年月,'" & Gfmt年月 & "') AS I_金利初回年月,"
    wstr = wstr + "FORMAT(K.解約実行日,'" & Gfmt年月日 & "') AS I_解約実行日,"
    wstr = wstr + "FORMAT(K.金融解約実行日,'" & Gfmt年月日 & "') AS I_金融解約実行日,"
'    wstr = wstr + "FORMAT(K.初回返済額,'#,###,###,###,##0円') AS I_初回返済額,"
'    wstr = wstr + "FORMAT(K.毎月返済額,'#,###,###,###,##0円') AS I_毎月返済額,"
'    wstr = wstr + "FORMAT(K.最終返済額,'#,###,###,###,##0円') AS I_最終返済額,"
    wstr = wstr & " IIF(K.手入力区分 = " & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",FORMAT(K.初回返済額,'#,###,###,###,##0円'),'**') As I_初回返済額,"
    wstr = wstr & " IIF(K.手入力区分 = " & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",FORMAT(K.毎月返済額,'#,###,###,###,##0円'),'**') As I_毎月返済額,"
    wstr = wstr & " IIF(K.手入力区分 = " & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",FORMAT(K.最終返済額,'#,###,###,###,##0円'),'**') As I_最終返済額,"
'    wstr = wstr + "FORMAT(K.返済単位月数,'##ヵ月') AS I_返済単位月数,"
    wstr = wstr & " IIF(K.手入力区分 = " & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",FORMAT(K.返済単位月数,'##ヵ月'),'**') As I_返済単位月数,"
    wstr = wstr + "K.担保名 AS I_担保名,"
    wstr = wstr + "K.資金用途 AS I_資金用途,"
    wstr = wstr + "K.金利条件 AS I_金利条件,"
    
    wstr = wstr + "KK.基準金利名 AS I_基準金利名,"
    wstr = wstr + "G.銀行名 AS I_銀行名,"
    wstr = wstr + "G.金融機関番号 AS I_金融機関番号,"
    wstr = wstr + "G.金融機関名 AS I_金融機関名,"
    wstr = wstr + "G.支店番号 AS I_支店番号,"
    wstr = wstr + "G.支店名 AS I_支店名,"
    wstr = wstr + "G.預金種別 AS I_預金種別,"
    wstr = wstr + "G.口座番号 AS I_口座番号,"
    wstr = wstr + "KS.借入金種別名 AS I_借入金種別名,"
    wstr = wstr + "B.部門番号 AS I_部門番号,"
    wstr = wstr + "B.部門名 AS I_部門名,"
    wstr = wstr + "B.部門略名 AS I_部門略名,"
    wstr = wstr + "KSM.金利グループ名 AS I_金利グループ名"
    
    wstr = wstr + " FROM ((((DBDA010_借入金 AS K"
    wstr = wstr + " LEFT JOIN DAAA040_銀行マスタ AS G"
    wstr = wstr + "  ON G.銀行番号 = K.銀行番号)"
    wstr = wstr + " LEFT JOIN DAAA116_借入金種別 AS KS"
    wstr = wstr + "  ON KS.借入金種別区分 = K.借入金種別区分)"
    wstr = wstr + " LEFT JOIN DAAA200_部門マスタ AS B"
    wstr = wstr + "  ON B.部門番号 = K.プロジェクト番号)"
    wstr = wstr + " LEFT JOIN DAAA115_金利シミュレーショングループ AS KSM"
    wstr = wstr + "  ON KSM.金利グループ区分 = K.金利グループ区分)"
    wstr = wstr + " LEFT JOIN DAAA116_基準金利 AS KK"
    wstr = wstr + "  ON KK.基準金利区分 = K.基準金利区分"
    
    wstr = wstr + " Where K.借入番号 = '" & P8.FCStr(GRpt.コンボ_01) + "'"

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
'
    ' =========================================
    '           　 CsvFile 作成
    ' =========================================
    If GRpt.CSV = 1 Then
        Call MX040_CsvOut_KARI
    End If

    ' =========================================
    '           　 ボタン制御
    ' =========================================

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
End Sub

'------------------------------------------------
' PageHeader_BeforePrint
'------------------------------------------------
Private Sub PageHeader_BeforePrint()
'
End Sub

'------------------------------------------------
' ReportFooter_BeforePrint
'------------------------------------------------
Private Sub ReportFooter_BeforePrint()
''

''
End Sub




