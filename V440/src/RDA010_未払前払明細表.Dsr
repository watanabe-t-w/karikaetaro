VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RDA010_未払前払明細表 
   Caption         =   "未払前払明細表"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   StartUpPosition =   3  'Windows の既定値
   WindowState     =   2  '最大化
   _ExtentX        =   21114
   _ExtentY        =   10636
   SectionData     =   "RDA010_未払前払明細表.dsx":0000
End
Attribute VB_Name = "RDA010_未払前払明細表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "RDA010_未払前払明細表"
'
Dim wd利息額1 As Double, wd利息額2 As Double
Dim wi日数1 As Integer, wi日数2 As Integer
Dim ws利率1 As Single, ws利率2 As Single
Dim ws番号1 As String, ws番号2 As String
Dim ws年月日1 As String, ws年月日2 As String
Dim ws開始日1 As String, ws開始日2 As String
Dim ws終了日1 As String, ws終了日2 As String
'
'------------------------------------------------
' ActiveReport_ReportStart
'------------------------------------------------
Private Sub ActiveReport_ReportStart()
'
    Dim w借入データ As MAA910_借入金
    
    Dim wRs As ADODB.Recordset
    Dim wstr As String, wWhere As String
    
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
'
    '----------------------------------------------------------------
    '                           ** 見出し **
    '----------------------------------------------------------------
    出力日 = Now
    企業名 = GCoName
    
    '----------------------------------------------------------------
    '          ** グループ ・ Where ・ OrderＢｙ **
    '----------------------------------------------------------------
    GroupHeader1.DataField = "GrpFld_D"
    GroupHeader2.DataField = "GrpFld_M"

    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    '** 明細ファイル 削除 **
    wstr = ""
    wstr = wstr + "Delete * From DCDA030_利息未払前払明細"
    GDb.Execute wstr
'
'    FLG_TR = False
'
'    wstr = ""
'    wstr = wstr + "Select *"
'    wstr = wstr + " From DBDA010_借入金 As k"
'    wstr = wstr + " Where K.借入番号 = '" & GRpt.コンボ_01 + "'"
'    wstr = wstr + wWhere
'    Call AdoRecordsetOpen(GDb, wRs, wstr)
'      Do Until wRs.EOF
'
'          w借入データ = MBD010_借入データセット(wRs)
'          'If P8.FCDbl(wRs("手入力区分")) = "0" Then
'          ''標準
'          '    Call MBD010_借入金テーブル作成(GRpt.金融, w借入データ)
'          'Else
'          ''入力登録
'          '    Call MBD010_借入金入力明細Read(w借入データ)
'          '
'          '    FLG_TR = True
'          'End If
'
'          wRs.MoveNext
'      Loop
'    wRs.Close
'    Set wRs = Nothing

    Call MBD020_借入金ワークテーブル作成("DBDA010_借入金", "")
    
    Call MRB010_標準入力未払前払("DCIA010_借入金ワーク", GRpt.コンボ_01)
    Call MRB010_手入力未払前払("DCIA010_借入金ワーク", GRpt.コンボ_01)

'
    '通常 or 手入力 の書式設定は
    H00_返済単位月数.Visible = True
    H00_支払区分.Visible = True
    H00_営業日.Visible = True
    H00_利息日数.Visible = True
    H00_利息支払方法.Visible = True
    H00_金利年間日数.Visible = True
    H00_据置回数.Visible = True
    
    If FLG_TR = True Then
        H00_返済単位月数.Visible = False
        H00_支払区分.Visible = False
        H00_営業日.Visible = False
        H00_利息日数.Visible = False
        H00_利息支払方法.Visible = False
        H00_据置回数.Visible = False
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
    
    wstr = wstr + " K.借入計画番号 As H00_借入計画番号,"
    wstr = wstr + " K.金融リストラ番号 As H00_金融リストラ番号,"
    wstr = wstr + " K.借入番号 As H00_借入番号,"
    wstr = wstr + " K.借入内容 As H00_借入内容,"
    
    wstr = wstr + " Format(K.実行日,'" & Gfmt年月日 & "') As H00_実行日,"
    wstr = wstr + " Format(K.初回返済年月,'" & Gfmt年月 & "') As H00_初回返済年月,"
    wstr = wstr + " Format(K.最終返済年月,'" & Gfmt年月 & "') As H00_最終返済年月,"
    wstr = wstr + " Format(K.解約実行日,'" & Gfmt年月日 & "') As H00_解約年月日,"
    wstr = wstr + " Format(K.金融解約実行日,'" & Gfmt年月日 & "') As H00_金融解約日,"
    wstr = wstr + " K.融資金額 As H00_融資金額,"
    
    '手入力の場合は借入金データセットしない
    'HederとDetailの所で表示制御
    If FLG_TR <> True Then
        
        wstr = wstr + " Format(K.支払回数,'#,##0') As H00_支払回数,"
        wstr = wstr + " Format(K.据置回数,'#,##0') As H00_据置回数,"
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
        
        wstr = wstr + " Format(" & w借入データ.支払回数 & ",'#,##0') As H00_支払回数,"
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
    wstr = wstr + " IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "最終返済日控除")) & ",'最終返済日控除','実行日及び最終返済日控除'))) As H00_利息控除区分,"
    wstr = wstr + " IIF(K.金利計算年間日数 = " & P8.FCDbl(XMXA020_区分("金利計算", "365日")) & ",'365日','360日') As H00_金利年間日数,"
    wstr = wstr + " IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As H00_金利種別,"
    wstr = wstr + " K.金利条件 As H00_金利条件,"
    wstr = wstr + " IIF(K.有担保フラグ=0,'無担保','有担保') As H00_担保区分,"
    wstr = wstr & " IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期借入金','長期借入金') AS H00_長短区分,"
    wstr = wstr + " IIF(K.設備フラグ=0,'運転資金','設備') As H00_設備区分,"
    
    wstr = wstr + " K.担保名 As H00_担保名,"
    wstr = wstr + " K.資金用途 As H00_資金用途,"
    wstr = wstr + " KS.借入金種別名 As H00_借入金種別名,"
    wstr = wstr + " KK.基準金利名 As H00_基準金利名,"
    wstr = wstr + " KG.金利グループ名 As H00_金利グループ名,"
    
    'wstr = wstr + " H.保証会社区分名 As H00_保証会社区分名,"
    'wstr = wstr + " K.保証料率 As H00_保証料率,"
    'wstr = wstr + " Y.融資区分名 As H00_融資区分名,"
    
    wstr = wstr & "Format(KM.返済年月日,'" & Gfmt年月 & "') As GrpFld_M,"
    wstr = wstr & "Format(KM.返済年月日,'" & Gfmt年月日 & "') As GrpFld_D,"
    wstr = wstr & "Format(KM.返済年月日,'" & Gfmt年月 & "') As I_年月,"
    wstr = wstr & "Format(KM.返済年月日,'" & Gfmt年月日 & "') As I_年月日,"
    wstr = wstr & "KM.月毎NO As I_番号,"
    wstr = wstr & "KM.利息額増 As I_利息額増,"
    wstr = wstr & "KM.利息額減 As I_利息額減,"
    wstr = wstr & "KM.日割日数 As I_日数,"
    wstr = wstr & "Format(KM.利率,'#,##0.00000') As I_利率,"
    wstr = wstr & "Format(KM.開始年月日,'" & Gfmt年月日 & "') As I_開始日,"
    wstr = wstr & "Format(KM.終了年月日,'" & Gfmt年月日 & "') As I_終了日"
    
    wstr = wstr + " From (((((((DCDA030_利息未払前払明細  As KM"
    wstr = wstr + " Inner Join DBDA010_借入金 As K"
    wstr = wstr + "  ON KM.借入番号 = K.借入番号)"
    wstr = wstr + " Inner Join DAAA040_銀行マスタ As G"
    wstr = wstr + "  ON K.銀行番号 = G.銀行番号)"
    wstr = wstr + " Inner Join DAAA116_借入金種別 As KS"
    wstr = wstr + "  ON K.借入金種別区分 = KS.借入金種別区分)"
    wstr = wstr + " Inner Join DAAB020_支払区分マスタ As S"
    wstr = wstr + "  ON K.支払日 = S.支払日)"
    wstr = wstr + " Left Join DAAA116_基準金利 As KK"
    wstr = wstr + "  ON K.基準金利区分 = KK.基準金利区分)"
    wstr = wstr + " Left Join DAAA115_金利シミュレーショングループ As KG"
    wstr = wstr + "  ON K.金利グループ区分 = KG.金利グループ区分)"
    wstr = wstr + " Left Join DAAA100_保証会社区分 As H"
    wstr = wstr + "  ON K.保証会社区分 = H.保証会社区分)"
    wstr = wstr + " Left Join DAAA110_融資区分 As Y"
    wstr = wstr + "  ON K.融資区分 = Y.融資区分"
    
    wstr = wstr & " Order BY KM.返済年月日,KM.月毎NO"
    
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
Resume
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
' PageHeader_BeforePrint
'------------------------------------------------
Private Sub PageHeader_BeforePrint()
'
    If Me.PageHeader.Controls("H00_利息区分") = "利息先払" Then
        Me.PageHeader.Controls("L_利息1").Caption = "前払利息減"
        Me.PageHeader.Controls("L_利息2").Caption = "前払利息増"
        Me.PageHeader.Controls("L_利息増").Caption = "前払利息増"
        Me.PageHeader.Controls("L_利息減").Caption = "前払利息減"
        Me.PageHeader.Controls("L_利息残高").Caption = "前払利息残高"
    ElseIf Me.PageHeader.Controls("H00_利息区分") = "利息後払" Then
        Me.PageHeader.Controls("L_利息1").Caption = "未払利息増"
        Me.PageHeader.Controls("L_利息2").Caption = "未払利息減"
        Me.PageHeader.Controls("L_利息増").Caption = "未払利息増"
        Me.PageHeader.Controls("L_利息減").Caption = "未払利息減"
        Me.PageHeader.Controls("L_利息残高").Caption = "未払利息残高"
    End If
'
End Sub

'------------------------------------------------
' Detail_BeforePrint
'------------------------------------------------
Private Sub Detail_BeforePrint()
'
    If P8.FCStr(Me.PageHeader.Controls("H00_利息区分")) = "利息先払" Then
    '利息先払 1：減、2：増
        If P8.FCDbl(Me.Detail.Controls("I_利息額減")) <> 0 Then
            ws番号1 = P8.FCStr(Me.Detail.Controls("I_番号"))
            ws年月日1 = P8.FCStr(Me.Detail.Controls("I_年月日"))
            wd利息額1 = P8.FCDbl(Me.Detail.Controls("I_利息額減"))
            ws開始日1 = P8.FCStr(Me.Detail.Controls("I_開始日"))
            ws終了日1 = P8.FCStr(Me.Detail.Controls("I_終了日"))
            wi日数1 = P8.FCDbl(Me.Detail.Controls("I_日数"))
            ws利率1 = P8.FCDbl(Me.Detail.Controls("I_利率"))
        End If
        
        If P8.FCDbl(Me.Detail.Controls("I_利息額増")) <> 0 Then
            ws番号2 = P8.FCStr(Me.Detail.Controls("I_番号"))
            ws年月日2 = P8.FCStr(Me.Detail.Controls("I_年月日"))
            wd利息額2 = P8.FCDbl(Me.Detail.Controls("I_利息額増"))
            ws開始日2 = P8.FCStr(Me.Detail.Controls("I_開始日"))
            ws終了日2 = P8.FCStr(Me.Detail.Controls("I_終了日"))
            wi日数2 = P8.FCDbl(Me.Detail.Controls("I_日数"))
            ws利率2 = P8.FCDbl(Me.Detail.Controls("I_利率"))
        End If
    End If
'
    If P8.FCStr(Me.PageHeader.Controls("H00_利息区分")) = "利息後払" Then
    '利息後払 1：増、2：減
        If P8.FCDbl(Me.Detail.Controls("I_利息額増")) <> 0 Then
            ws番号1 = P8.FCStr(Me.Detail.Controls("I_番号"))
            ws年月日1 = P8.FCStr(Me.Detail.Controls("I_年月日"))
            wd利息額1 = P8.FCDbl(Me.Detail.Controls("I_利息額増"))
            ws開始日1 = P8.FCStr(Me.Detail.Controls("I_開始日"))
            ws終了日1 = P8.FCStr(Me.Detail.Controls("I_終了日"))
            wi日数1 = P8.FCDbl(Me.Detail.Controls("I_日数"))
            ws利率1 = P8.FCDbl(Me.Detail.Controls("I_利率"))
        End If
        
        If P8.FCDbl(Me.Detail.Controls("I_利息額減")) <> 0 Then
            ws番号2 = P8.FCStr(Me.Detail.Controls("I_番号"))
            ws年月日2 = P8.FCStr(Me.Detail.Controls("I_年月日"))
            wd利息額2 = P8.FCDbl(Me.Detail.Controls("I_利息額減"))
            ws開始日2 = P8.FCStr(Me.Detail.Controls("I_開始日"))
            ws終了日2 = P8.FCStr(Me.Detail.Controls("I_終了日"))
            wi日数2 = P8.FCDbl(Me.Detail.Controls("I_日数"))
            ws利率2 = P8.FCDbl(Me.Detail.Controls("I_利率"))
        End If
    End If
'
End Sub

'------------------------------------------------
' GroupFooter1_BeforePrint
'------------------------------------------------
Private Sub GroupFooter1_BeforePrint()
'
    Dim ws金利年間日数 As String
'
    If P8.FCStr(Me.PageHeader.Controls("H00_金利年間日数")) = "365日" Then
        ws金利年間日数 = "365"
    ElseIf P8.FCStr(Me.PageHeader.Controls("H00_金利年間日数")) = "360日" Then
        ws金利年間日数 = "360"
    End If
'
    If wd利息額1 <> 0 Then
        Me.GroupFooter1.Controls("G10_番号1").BackStyle = 1
        Me.GroupFooter1.Controls("G10_番号1") = ws番号1
        Me.GroupFooter1.Controls("G10_年月日1") = ws年月日1
        Me.GroupFooter1.Controls("G10_利息額1") = Format(wd利息額1, "#,##0")
        Me.GroupFooter1.Controls("G10_期間1") = ws開始日1 & "～" & ws終了日1
        Me.GroupFooter1.Controls("G10_日数1") = wi日数1
        Me.GroupFooter1.Controls("G10_利率1") = Format(ws利率1, "#,##0.00000")
        Me.GroupFooter1.Controls("G10_式1") = "= " & Me.GroupFooter1.Controls("G10_利息額1") & " × " & Me.GroupFooter1.Controls("G10_利率1") & " × " & Me.GroupFooter1.Controls("G10_日数1") & " / " & ws金利年間日数
    Else
        Me.GroupFooter1.Controls("G10_番号1").BackStyle = 0
        Me.GroupFooter1.Controls("G10_番号1") = ""
        Me.GroupFooter1.Controls("G10_年月日1") = ""
        Me.GroupFooter1.Controls("G10_利息額1") = ""
        Me.GroupFooter1.Controls("G10_期間1") = ""
        Me.GroupFooter1.Controls("G10_日数1") = ""
        Me.GroupFooter1.Controls("G10_利率1") = ""
        Me.GroupFooter1.Controls("G10_式1") = ""
    End If
'
    If wd利息額2 <> 0 Then
        Me.GroupFooter1.Controls("G10_番号2").BackStyle = 1
        Me.GroupFooter1.Controls("G10_番号2") = ws番号2
        Me.GroupFooter1.Controls("G10_年月日2") = ws年月日2
        Me.GroupFooter1.Controls("G10_利息額2") = Format(wd利息額2, "#,##0")
        Me.GroupFooter1.Controls("G10_期間2") = ws開始日2 & "～" & ws終了日2
        Me.GroupFooter1.Controls("G10_日数2") = wi日数2
        Me.GroupFooter1.Controls("G10_利率2") = Format(ws利率2, "#,##0.00000")
        Me.GroupFooter1.Controls("G10_式2") = "= " & Me.GroupFooter1.Controls("G10_利息額2") & " × " & Me.GroupFooter1.Controls("G10_利率2") & " × " & Me.GroupFooter1.Controls("G10_日数2") & " / " & ws金利年間日数
    Else
        Me.GroupFooter1.Controls("G10_番号2").BackStyle = 0
        Me.GroupFooter1.Controls("G10_番号2") = ""
        Me.GroupFooter1.Controls("G10_年月日2") = ""
        Me.GroupFooter1.Controls("G10_利息額2") = ""
        Me.GroupFooter1.Controls("G10_期間2") = ""
        Me.GroupFooter1.Controls("G10_日数2") = ""
        Me.GroupFooter1.Controls("G10_利率2") = ""
        Me.GroupFooter1.Controls("G10_式2") = ""
    End If
'
End Sub

'------------------------------------------------
' GroupFooter1_AfterPrint
'------------------------------------------------
Private Sub GroupFooter1_AfterPrint()
'
    ws番号1 = ""
    ws年月日1 = ""
    wd利息額1 = 0
    ws開始日1 = ""
    ws終了日1 = ""
    wi日数1 = 0
    ws利率1 = 0
    
    ws番号2 = ""
    ws年月日2 = ""
    wd利息額2 = 0
    ws開始日2 = ""
    ws終了日2 = ""
    wi日数2 = 0
    ws利率2 = 0
'
End Sub

'------------------------------------------------
' GroupFooter2_BeforePrint
'------------------------------------------------
Private Sub GroupFooter2_BeforePrint()
'
    Dim wd01 As Double
'
    Me.GroupFooter2.Controls("G20_利息額1") = ""
    Me.GroupFooter2.Controls("G20_利息額2") = ""
'
    If P8.FCStr(Me.PageHeader.Controls("H00_利息区分")) = "利息先払" Then
    '利息先払 1：減、2：増
        Me.GroupFooter2.Controls("G20_利息額1") = Format(Me.GroupFooter2.Controls("G20_利息額減"), "#,##0")
        Me.GroupFooter2.Controls("G20_利息額2") = Format(Me.GroupFooter2.Controls("G20_利息額増"), "#,##0")
    End If
'
    If P8.FCStr(Me.PageHeader.Controls("H00_利息区分")) = "利息後払" Then
    '利息後払 1：増、2：減
        Me.GroupFooter2.Controls("G20_利息額1") = Format(Me.GroupFooter2.Controls("G20_利息額増"), "#,##0")
        Me.GroupFooter2.Controls("G20_利息額2") = Format(Me.GroupFooter2.Controls("G20_利息額減"), "#,##0")
    End If
'
    Me.GroupFooter2.Controls("G20_利息額増") = Format(Me.GroupFooter2.Controls("G20_利息額増"), "#,##0")
    Me.GroupFooter2.Controls("G20_利息額減") = Format(Me.GroupFooter2.Controls("G20_利息額減"), "#,##0")
    
    wd01 = P8.FCDbl(Me.GroupFooter2.Controls("G20_利息額増")) - P8.FCDbl(Me.GroupFooter2.Controls("G20_利息額減"))
    If wd01 < 0 Then
        wd01 = 0
    End If
    Me.GroupFooter2.Controls("G20_融資残高") = Format(wd01, "#,##0")
'
End Sub
