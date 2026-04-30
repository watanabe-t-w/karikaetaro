VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} REA010_借入金時価評価明細表 
   Caption         =   "借入金時価評価明細表"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12855
   Icon            =   "REA010_借入金時価評価明細表.dsx":0000
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   _ExtentX        =   22675
   _ExtentY        =   14473
   SectionData     =   "REA010_借入金時価評価明細表.dsx":0ECA
End
Attribute VB_Name = "REA010_借入金時価評価明細表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "REA010_借入金時価評価明細表"

Dim wstr As String, w円単位 As String
Dim w分母 As Integer
Dim FLG_TR As Boolean
Dim wd決算日1 As Date, wd決算日2 As Date
Dim wd元金額(1) As Double
Dim wd現在価値(1) As Double, wdG現在価値 As Double
'
'------------------------------------------------
' ActiveReport_ReportStart
'------------------------------------------------
Private Sub ActiveReport_ReportStart()
'
    Dim wRs As ADODB.Recordset
    Dim wWhere As String
    
    Dim w借入データ As MAA910_借入金
    Dim w時価一覧 As MAA910_時価評価一覧
    
    Dim wiH支払回数 As Integer, wiH据置回数 As Integer
    Dim wD実行年月 As Date
    
    Dim wd期末残高 As Double
    Dim wiCnt As Integer
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
    Printer.Orientation = ddOPortrait
'
    Me.Detail.Controls("S_Line").Visible = False
'
    '金剛石 or 借換たろう！
    Me.PageHeader.Controls("L_計画番号").Visible = True
    Me.PageHeader.Controls("H00_借入計画番号").Visible = True
    If GProduct <> "金剛石" Then
        Me.PageHeader.Controls("L_計画番号").Visible = False
        Me.PageHeader.Controls("H00_借入計画番号").Visible = False
    End If
'
    Me.PageHeader.Controls("L_帳票名").Caption = "借入金時価評価明細表"
    Me.PageHeader.Controls("L_番号").Caption = "借入番号"
    Me.PageHeader.Controls("L_内容").Caption = "借入内容"
'
    '----------------------------------------------------------------
    '                           ** 見出し **
    '----------------------------------------------------------------
    出力日 = Now
    企業名 = GCoName
    
    wd決算日1 = P8.FCDate(GRpt.テキスト_01)
    wd決算日2 = C年月日.GetDate("設定", DateAdd("yyyy", 1, wd決算日1), G基本情報.決算締日)
    
    If GRpt.千円単位 = 1 Then
        w分母 = 1000
        L_単位.Caption = "（千円単位）"
        w円単位 = "千円"
    Else
        w分母 = 1
        L_単位 = "（円単位）"
        w円単位 = "円"
    End If
    
    '----------------------------------------------------------------
    '          ** グループ ・ Where ・ OrderＢｙ **
    '----------------------------------------------------------------
    wWhere = ""
    
    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    wiCnt = 0
    FLG_TR = False
    
    'GRet = MAA200_時価評価適用金利作成(wd決算日1)
    
    '** 明細ファイル 削除 **
    wstr = ""
    wstr = wstr + "Delete * From DCDA010_借入金時価評価明細"
    GDb.Execute wstr
    
    '** 明細ファイル 作成 **
    wstr = ""
    wstr = wstr + "Select *"
    wstr = wstr + " From DBDA010_借入金 As k"
    wstr = wstr + " Where K.借入番号 = '" & GRpt.コンボ_01 + "'"
'    wstr = wstr & " and K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "固定金利")) DEL 20170501 M.Mino
    wstr = wstr & " and K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "長期借入金"))
    wstr = wstr & " and Format(K.実行日,'yyyymmdd') <= '" & Format(wd決算日1, "yyyymmdd") & "'"
    wstr = wstr & " and Format(K.最終返済実行日,'yyyymmdd') > '" & Format(wd決算日1, "yyyymmdd") & "'"
    wstr = wstr & " and K.手入力区分 <> 2 And K.取消フラグ = 0"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.EOF Then
        ReDim G時価明細(0)
    Else
        Do Until wRs.EOF
        
            w借入データ = MBD010_借入データセット(wRs)
            
            If wiCnt = 0 Then
                'GRet = MAA200_時価評価適用金利作成(wd決算日1, w借入データ.基準金利区分) DEL 20170501 M.Mino
                GRet = MAA200_時価評価適用金利作成(wd決算日1, w借入データ.基準金利区分, P8.FCStr(XMXA020_区分("金利種別", wRs("金利種別")))) 'ADD 20170501 M.Mino
                wiCnt = wiCnt + 1
            End If
            
            If P8.FCDbl(wRs("手入力区分")) = "0" Then
            '標準
                Call MBD010_借入金テーブル作成("", w借入データ)
                w時価一覧 = MAA200_時価評価明細作成(w借入データ, wd決算日1)
            Else
            '入力登録
                Call MBD010_借入金入力明細Read(w借入データ)
                w時価一覧 = MAA200_時価評価明細作成_入力登録(w借入データ, wd決算日1)
                
                FLG_TR = True
            End If
            
              '支払回数 据置回数 再計算 2014/08/19
              wiH支払回数 = w借入データ.支払回数
              wiH据置回数 = w借入データ.据置回数
              If P8.FCDbl(wRs("手入力区分")) = "0" Then
                  If w借入データ.返済単位月数 = 1 Then
                  '一括返済
                      If w借入データ.初回返済額 = 0 And w借入データ.毎月返済額 = 0 And w借入データ.最終返済額 = w借入データ.融資金額 Then
                          If CInt(Right(w借入データ.実行日, 2)) >= w借入データ.支払日 Then
                              wD実行年月 = DateAdd("m", 1, w借入データ.実行日)
                          ElseIf CInt(Right(w借入データ.実行日, 2)) <= w借入データ.支払日 Then
                              wD実行年月 = Format(w借入データ.実行日, "yyyy/mm/01")
                          End If
                          
                          wiH支払回数 = 1
                          wiH据置回数 = DateDiff("m", wD実行年月, w借入データ.最終返済年月)
                      
                      End If
              
                  ElseIf w借入データ.返済単位月数 > 1 Then
                      If w借入データ.初回返済額 = 0 And w借入データ.毎月返済額 = 0 And w借入データ.最終返済額 = w借入データ.融資金額 Then
                          If CInt(Right(w借入データ.実行日, 2)) >= w借入データ.支払日 Then
                              wD実行年月 = DateAdd("m", 1, w借入データ.実行日)
                          ElseIf CInt(Right(w借入データ.実行日, 2)) <= w借入データ.支払日 Then
                              wD実行年月 = Format(w借入データ.実行日, "yyyy/mm/01")
                          End If
                          
                          wiH支払回数 = 1
                          wiH据置回数 = DateDiff("m", wD実行年月, w借入データ.最終返済年月)
                      
                      Else
                          wiH支払回数 = P8.FFix(P8.FCDiv(w借入データ.支払回数, w借入データ.返済単位月数) + 1)
                      End If
                  End If
              Else
                  wiH支払回数 = 0
                  wiH据置回数 = 0
              End If
              
              If wiH据置回数 < 0 Then
                  wiH据置回数 = 0
              End If
            
            wRs.MoveNext
        Loop
    End If
    
    wRs.Close
    Set wRs = Nothing
'
    wd期末残高 = MAA200_時価評価明細
'
    '----------------------------------------------------------------
    '                          ** 合計 計算 **
    '----------------------------------------------------------------
    Me.PageHeader.Controls("H00_期末残高") = Format(P8.FCDblRD(wd期末残高) / w分母, "#,##0" + w円単位)
    '----------------------------------------------------------------
    '                       ** レコード　ソース **
    '----------------------------------------------------------------
    wstr = ""
    wstr = wstr + "Select "
    wstr = wstr + " K.借入計画番号 As H00_借入計画番号,"
'    wstr = wstr + " K.金融リストラ番号 As H00_金融リストラ番号,"
    wstr = wstr + " K.借入番号 As H00_借入番号,"
    wstr = wstr + " K.借入内容 As H00_借入内容,"
    
    wstr = wstr + " Format(K.実行日,'" & Gfmt年月日 & "') As H00_実行日,"
    wstr = wstr + " Format(K.初回返済年月,'" & Gfmt年月 & "') As H00_初回返済年月,"
    wstr = wstr + " Format(K.最終返済年月,'" & Gfmt年月 & "') As H00_最終返済年月,"
    wstr = wstr + " Format(K.解約実行日,'" & Gfmt年月日 & "') As H00_解約年月日,"
'    wstr = wstr + " Format(K.金融解約実行日,'" & Gfmt年月日 & "') As H00_金融解約日,"
    wstr = wstr + " K.融資金額 As H00_融資金額,"
    
    If FLG_TR <> True Then
        '支払回数 据置回数 2014/08/19
        wstr = wstr + " Format(" & wiH支払回数 & ",'#,##0') As H00_支払回数,"
        wstr = wstr + " Format(" & wiH据置回数 & ",'#,##0') As H00_据置回数,"
        
        wstr = wstr + " K.返済単位月数 As H00_返済単位月数,"
        
        '変動金利の場合
        If P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) = w借入データ.金利種別 Then
            If w借入データ.変動最終利率 > -1 Then
                wstr = wstr + " '*' As H00_利率フラグ,"
                wstr = wstr + "'" & w借入データ.変動最終利率 & "' As H00_利率,"
            Else
                wstr = wstr + " '' As H00_利率フラグ,"
                wstr = wstr + " K.利率 As H00_利率,"
            End If
        Else
            wstr = wstr + " K.利率 As H00_利率,"
            wstr = wstr + " '' As H00_利率フラグ,"
        End If
    Else
        
        '支払回数 据置回数 2014/08/19
        wstr = wstr + " '' As H00_返済単位月数,"
        wstr = wstr + " Format(" & wiH支払回数 & ",'#,##0') As H00_支払回数,"
        wstr = wstr + " Format(" & wiH据置回数 & ",'#,##0') As H00_据置回数,"
        
        If P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) = w借入データ.金利種別 Then
            If w借入データ.変動最終利率 > -1 Then
                wstr = wstr + " '*' As H00_利率フラグ,"
                wstr = wstr + "'" & w借入データ.変動最終利率 & "' As H00_利率,"
            Else
                wstr = wstr + " '' As H00_利率フラグ,"
                wstr = wstr + " K.利率 As H00_利率,"
            End If
        Else
            wstr = wstr + " K.利率 As H00_利率,"
            wstr = wstr + " '' As H00_利率フラグ,"
        End If
    End If
    
    wstr = wstr + " G.銀行名 As H00_銀行名,"
    wstr = wstr + " K.支払日,"
    wstr = wstr + " S.支払区分名 As H00_支払区分,"
    wstr = wstr + " IIF(K.営業日区分 = " & P8.FCDbl(XMXA020_区分("営業日", "翌営業日")) & ",'翌営業日','前営業日') As H00_営業日,"
    wstr = wstr + " IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As H00_利息区分,"
    wstr = wstr + " IIF(K.利息計算日数区分 = " & P8.FCDbl(XMXA020_区分("利息計算日数", "営業日数")) & ",'営業日数','固定日数') As H00_利息日数,"
    wstr = wstr + " IIF(K.利息支払方法 = " & P8.FCDbl(XMXA020_区分("利息支払", "毎月")) & ",'毎月','一括') As H00_利息支払方法,"
    wstr = wstr + " IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "控除無し")) & ",'控除無し',"
    wstr = wstr + " IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "実行日控除")) & ",'実行日控除',"
    wstr = wstr + " IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "最終返済日控除")) & ",'最終返済日控除',"
    wstr = wstr + " IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) & ",'実行日及び最終返済日控除','中間利払最終日控除')))) As H00_利息控除区分,"
    wstr = wstr + " IIF(K.金利計算年間日数 = " & P8.FCDbl(XMXA020_区分("金利計算", "365日")) & ",'365日','360日') As H00_金利年間日数,"
    wstr = wstr + " IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As H00_金利種別,"
    wstr = wstr + " K.金利条件 As H00_金利条件,"
    wstr = wstr + " IIF(K.有担保フラグ=0,'無担保','有担保') As H00_担保区分,"
    wstr = wstr & " IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期借入金','長期借入金') AS H00_長短区分,"
    wstr = wstr + " IIF(K.設備フラグ=0,'運転資金','設備') As H00_設備区分,"
    wstr = wstr + " IIF(K.手入力区分 = " & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",'標準登録','入力登録') As H00_登録方法,"
    
    wstr = wstr + " K.担保名 As H00_担保名,"
    wstr = wstr + " K.資金用途 As H00_資金用途,"
    wstr = wstr + " B.部門名 As H00_部門名,"
    wstr = wstr + " KS.借入金種別名 As H00_借入金種別名,"
    wstr = wstr + " KK.基準金利名 As H00_基準金利名,"
'    wstr = wstr + " KG.金利グループ名 As H00_金利グループ名,"
    
'    wstr = wstr + " Format(1,'#,##0') As I_回数,"
    wstr = wstr + " KM.時価評価金利 As H00_時価評価金利,"
    wstr = wstr + " Format(KM.決算年月日,'" & Gfmt年月日 & "') As H00_決算日,"
    wstr = wstr + " Format(KM.日割日数,'#,##0') As I_日割日数,"
    wstr = wstr + " Format(KM.返済年月日,'" & Gfmt年月日 & "') As I_返済年月日,"
    wstr = wstr + " Format(KM.利息計算年月日,'" & Gfmt年月日 & "') As I_利息計算年月日,"
    wstr = wstr + " KM.返済金額 As I_返済金額,"
    wstr = wstr + " KM.元金額 As I_元金額,"
    wstr = wstr + " KM.利息額 As I_利息額,"
    wstr = wstr + " KM.融資残高 As I_融資残高,"
    wstr = wstr + " KM.指数 As I_指数,"
    wstr = wstr + " KM.分母 As I_分母,"
    wstr = wstr + " KM.現価係数 As I_現価係数,"
    wstr = wstr + " KM.現在価値 As I_現在価値"

    wstr = wstr + " From ((((((DCDA010_借入金時価評価明細  As KM"
    wstr = wstr + " Inner Join DBDA010_借入金 As K"
    wstr = wstr + "  ON KM.借入番号 = K.借入番号)"
    wstr = wstr + " Inner Join DAAA040_銀行マスタ As G"
    wstr = wstr + "  ON K.銀行番号 = G.銀行番号)"
    wstr = wstr + " Inner Join DAAA116_借入金種別 As KS"
    wstr = wstr + "  ON K.借入金種別区分 = KS.借入金種別区分)"
    wstr = wstr + " Inner Join DAAB020_支払区分マスタ As S"
    wstr = wstr + "  ON K.支払日 = S.支払日)"
    wstr = wstr + " LEFT JOIN DAAA200_部門マスタ AS B"
    wstr = wstr + "  ON B.部門番号 = K.プロジェクト番号)"
    wstr = wstr + " Left Join DAAA116_基準金利 As KK"
    wstr = wstr + "  ON K.基準金利区分 = KK.基準金利区分)"
    wstr = wstr + " Left Join DAAA115_金利シミュレーショングループ As KG"
    wstr = wstr + "  ON K.金利グループ区分 = KG.金利グループ区分"
    
    wstr = wstr + " Order BY KM.返済年月日"
    
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

'------------------------------------------------
' Detail_BeforePrint
'------------------------------------------------
Private Sub Detail_BeforePrint()
'
    Dim ws01 As String
    Dim wiKubun As Integer
    Dim wv返済年月日 As Variant
    Dim wdGankin As Double, wdGenzaikachi As Double
'
    wv返済年月日 = C年月日.平成To西暦("年月日", P8.FCStr(Me.Detail.Controls("I_返済年月日")))
    If wd決算日1 < P8.FCDate(wv返済年月日) And P8.FCDate(wv返済年月日) <= wd決算日2 Then
        wiKubun = 0
    ElseIf wd決算日2 < P8.FCDate(wv返済年月日) Then
        wiKubun = 1
    End If
'
    Call MXA030_ReportColor(Me.Detail.Controls("I_元金額"))
    Call MXA030_ReportColor(Me.Detail.Controls("I_利息額"))
    Call MXA030_ReportColor(Me.Detail.Controls("I_返済金額"))
    Call MXA030_ReportColor(Me.Detail.Controls("I_融資残高"))
    Call MXA030_ReportColor(Me.Detail.Controls("I_日割日数"))
'    Call MXA030_ReportColor(Me.Detail.Controls("I_指数"))
'    Call MXA030_ReportColor(Me.Detail.Controls("I_分母"))
    Call MXA030_ReportColor(Me.Detail.Controls("I_現価係数"))
    Call MXA030_ReportColor(Me.Detail.Controls("I_現在価値"))
    
    wdGankin = P8.FCDbl(Me.Detail.Controls("I_元金額"))
    wdGenzaikachi = P8.FCDbl(Me.Detail.Controls("I_現在価値"))
    
    Me.Detail.Controls("I_元金額") = Format(P8.FCDblRD(Me.Detail.Controls("I_元金額")) / w分母, "#,##0")
    Me.Detail.Controls("I_利息額") = Format(P8.FCDblRD(Me.Detail.Controls("I_利息額")) / w分母, "#,##0")
    Me.Detail.Controls("I_返済金額") = Format(P8.FCDblRD(Me.Detail.Controls("I_返済金額")) / w分母, "#,##0")
    Me.Detail.Controls("I_融資残高") = Format(P8.FCDblRD(Me.Detail.Controls("I_融資残高")) / w分母, "#,##0")
'    Me.Detail.Controls("I_指数") = Format(P8.FCDblRD6(Me.Detail.Controls("I_指数")), "#,##0.000000")
'    Me.Detail.Controls("I_分母") = Format(P8.FCDblRD6(Me.Detail.Controls("I_分母")), "#,##0.000000")
    Me.Detail.Controls("I_現価係数") = Format(P8.FCDblRD6(Me.Detail.Controls("I_現価係数")), "#,##0.000000")
    Me.Detail.Controls("I_現在価値") = Format(P8.FCDblRD(Me.Detail.Controls("I_現在価値")) / w分母, "#,##0")
'
    wd元金額(wiKubun) = wd元金額(wiKubun) + wdGankin
    wd現在価値(wiKubun) = wd現在価値(wiKubun) + wdGenzaikachi
    wdG現在価値 = wdG現在価値 + wdGenzaikachi
'
End Sub

'------------------------------------------------
' PageHeader_BeforePrint
'------------------------------------------------
Private Sub PageHeader_BeforePrint()
'
    Dim wd01 As Double
'
    Call MXA030_ReportColor(Me.PageHeader.Controls("H00_融資金額"))
    Call MXA030_ReportColor(Me.PageHeader.Controls("H00_利率"))
    
    Me.PageHeader.Controls("H00_融資金額") = Format(P8.FCDblRD(Me.PageHeader.Controls("H00_融資金額")) / w分母, "#,##0" + w円単位)
    Me.PageHeader.Controls("H00_利率") = Format(P8.FCDblRD5(Me.PageHeader.Controls("H00_利率")) / 100, "#,##0.00000%")
    Me.PageHeader.Controls("H00_時価評価金利") = Format(P8.FCDblRD5(Me.PageHeader.Controls("H00_時価評価金利")) / 100, "#,##0.00000%")
'
End Sub

'------------------------------------------------
' ReportFooter_BeforePrint
'------------------------------------------------
Private Sub ReportFooter_BeforePrint()
'
    Dim wdGankin As Double
'
    wdGankin = P8.FCDbl(Me.ReportFooter.Controls("G92_元金額"))

    Call MXA030_ReportColor(Me.ReportFooter.Controls("G90_元金額"))
    Call MXA030_ReportColor(Me.ReportFooter.Controls("G91_元金額"))
    Call MXA030_ReportColor(Me.ReportFooter.Controls("G92_元金額"))
    Call MXA030_ReportColor(Me.ReportFooter.Controls("G90_現在価値"))
    Call MXA030_ReportColor(Me.ReportFooter.Controls("G91_現在価値"))
    Call MXA030_ReportColor(Me.ReportFooter.Controls("G92_現在価値"))
    Call MXA030_ReportColor(Me.ReportFooter.Controls("G90_差額"))
    Call MXA030_ReportColor(Me.ReportFooter.Controls("G91_差額"))
    Call MXA030_ReportColor(Me.ReportFooter.Controls("G92_差額"))
    
    Me.ReportFooter.Controls("G90_元金額") = Format(wd元金額(0) / w分母, "#,##0")
    Me.ReportFooter.Controls("G91_元金額") = Format(wd元金額(1) / w分母, "#,##0")
    Me.ReportFooter.Controls("G92_元金額") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G92_元金額")) / w分母, "#,##0")
    
    Me.ReportFooter.Controls("G90_現在価値") = Format(wd現在価値(0) / w分母, "#,##0")
    Me.ReportFooter.Controls("G91_現在価値") = Format(wd現在価値(1) / w分母, "#,##0")
    Me.ReportFooter.Controls("G92_現在価値") = Format(wdG現在価値 / w分母, "#,##0")
    
    Me.ReportFooter.Controls("G90_差額") = Format((wd現在価値(0) - wd元金額(0)) / w分母, "#,##0")
    Me.ReportFooter.Controls("G91_差額") = Format((wd現在価値(1) - wd元金額(1)) / w分母, "#,##0")
    Me.ReportFooter.Controls("G92_差額") = Format((wdG現在価値 - wdGankin) / w分母, "#,##0")

'
End Sub




