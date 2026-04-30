VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RDF010_利息前払未払残高表 
   Caption         =   "利息未払前払残高表"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9555
   StartUpPosition =   3  'Windows の既定値
   WindowState     =   2  '最大化
   _ExtentX        =   16854
   _ExtentY        =   13070
   SectionData     =   "RDF010_利息未払前払残高表.dsx":0000
End
Attribute VB_Name = "RDF010_利息前払未払残高表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "RDF010_利息前払未払残高表"
'
Dim wML As Integer
Dim w番号 As String, wsTbl As String, wsTbl2 As String
Dim w分母 As Integer
Dim w推移表タイトル As MAA910_推移表タイトル
'
'------------------------------------------------
' ActiveReport_ReportStart
'------------------------------------------------
Private Sub ActiveReport_ReportStart()
'
    Dim wRs As ADODB.Recordset
    Dim wWhere As String
    
    Dim wstr As String
    Dim wsRet As String
    Dim ws_Ginko As String
    Dim j As Integer, k As Integer, l As Integer, wIndex As Integer
    
    Dim wdate As Date
    Dim w開始年月日 As Date
    Dim w推移表区分 As String, wsNengetu As String
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

    wML = 12 '列数
'
    'パラメータ部分
    Me.PageHeader.Controls("L_金融リストラ番号").Visible = True
    Me.PageHeader.Controls("H00_金融リストラ番号").Visible = True
'
    If GRpt.推移 = "年次" Then
        GRpt.テキスト_01 = GRpt.テキスト_02
        
        If G金利SM = True Then
            L_帳票名.Caption = " " & GRpt.帳票名 & " " & GRpt.テキスト_01 & "年度 - 金利SM " & GRpt.指定 & " " & GRpt.推移 & "- "
        Else
            L_帳票名.Caption = " " & GRpt.帳票名 & " " & GRpt.テキスト_01 & "年度 -" & GRpt.指定 & " " & GRpt.推移 & "- "
        End If
    Else
        GRpt.テキスト_02 = GRpt.テキスト_01
        
        wdate = C年月日.平成To西暦("年月日", GRpt.テキスト_02)
        GRpt.テキスト_01 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "e")
        wsNengetu = GRpt.テキスト_01
        
        If G金利SM = True Then
            L_帳票名.Caption = " " & GRpt.帳票名 & " " & GRpt.テキスト_02 & " - 金利SM " & GRpt.指定 & " " & GRpt.推移 & "- "
        Else
            L_帳票名.Caption = " " & GRpt.帳票名 & " " & GRpt.テキスト_02 & " -" & GRpt.指定 & GRpt.推移 & "- "
        End If
    End If
'
    '----------------------------------------------------------------
    '                           ** 見出し **
    '----------------------------------------------------------------
    出力日 = Now
    企業名 = GCoName
    
    H00_金融リストラ番号 = GRpt.金融
    
    If GRpt.千円単位 = 1 Then
        w分母 = 1000
        L_単位.Caption = "（千円単位）"
    Else
        w分母 = 1
        L_単位 = "（円単位）"
    End If
    
    '----------------------------------------------------------------
    '          ** グループ ・ Where ・ OrderＢｙ **
    '----------------------------------------------------------------
    
    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    Call MBD020_借入金ワークテーブル作成(wsTbl, GRpt.指定)
    Call MRB010_標準入力借入残高表固定日数("DCIA010_借入金ワーク")
    Call MRB010_手入力借入残高表("DCIA010_借入金ワーク")
'
    w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "平成")
    w推移表タイトル = MUA010_推移表ファイル作成("", "", w開始年月日, GRpt.推移, wML)

    wIndex = 1
    If GRpt.推移 <> "年次" Then
        For j = 1 To wML
            If GRpt.テキスト_02 = w推移表タイトル.X番目年月(j) Then
                    wIndex = j
                Exit For
            End If
        Next
    End If
    
    'CSVファイル パラメータ:GInt1
    GInt1 = wIndex
'
    '----------------------------------------------------------------
    '                          ** 合計 計算 **
    '----------------------------------------------------------------
    '印刷設定
    'Me.GroupFooter1.NewPage = ddNPAfter
    'Me.GroupFooter2.NewPage = ddNPAfter
    'Me.GroupFooter3.NewPage = ddNPAfter
    
    '印刷設定
    If GRpt.詳細表示 = 1 Then
        Me.Detail.Height = 210
    Else
        Me.Detail.Height = 0
    End If

'    If GRpt.指定 <> "" Then
'        Me.ReportFooter.Visible = False
'    Else
'        Me.ReportFooter.Visible = True
'    End If
'
    'グループセット
    'グループセット
    GroupHeader1.DataField = "GrpFld_Ginko"
    GroupHeader2.DataField = "GrpFld_RIsoku"
    If GStr <> "金利GR" Then
        GroupHeader3.DataField = "GrpFld_KShubetu"
    Else
        '金利SM
        GroupHeader3.DataField = "GrpFld_KGroup"
    
        GRet = 金利GR_CHECK
        If GRet <= 1 Then
            GroupHeader3.DataField = ""
            Me.GroupFooter3.Visible = False
        End If
    End If
'
    '----------------------------------------------------------------
    '                       ** レコード　ソース **
    '----------------------------------------------------------------
    w番号 = Right("00" + CStr(wIndex), 2)
    
    '** レコード　ソース **
    wstr = ""
    wstr = wstr & "Select "
    
    wstr = wstr & "K.借入金種別区分 As GrpFld_KShubetu,"
    wstr = wstr & "K.利息区分 As GrpFld_RIsoku,"
    wstr = wstr & "K.銀行番号 As GrpFld_Ginko,"
    
    If GStr <> "金利GR" Then
    wstr = wstr & "S.借入金種別名 As I_借入金種別名,"
        wstr = wstr & "S.借入金種別名 As G30_計名,"
    Else
        '金利SM
        wstr = wstr & "IIF(KG.金利グループ名<>'',KG.金利グループ名,'グループ無') As G30_計名,"
    End If
    
    wstr = wstr & "G.銀行名 As I_銀行名,"
    wstr = wstr & "K.借入番号 As I_借入番号,"
    wstr = wstr & "K.借入内容 As I_借入内容,"
    wstr = wstr + "IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As I_金利種別,"
    wstr = wstr + "KK.基準金利名 As I_基準金利名,"
    wstr = wstr & "K.金利条件 As I_金利備考,"
    wstr = wstr & "format(K.利率,'#,##0.00000') As I_利率,"
    wstr = wstr & " IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As I_利息区分,"
    
    wstr = wstr & "Z.残高_" & w番号 & " As I_融資残高,"
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & "-Z.前払利息増_" & w番号 & "+Z.前払利息_" & w番号 & ",-Z.未払利息増_" & w番号 & "+Z.未払利息減_" & w番号 & "+Z.未払利息_" & w番号 & ") As I_前月利息残高,"
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息増_" & w番号 & ",Z.未払利息増_" & w番号 & ") As I_利息増,"
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & ",Z.未払利息減_" & w番号 & ") As I_利息減,"
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息_" & w番号 & ",Z.未払利息_" & w番号 & ") As I_利息残高"
    
    If GStr <> "金利GR" Then
        wstr = wstr & " FROM (((DCDA010_借入残高推移表結果 As Z"
        wstr = wstr & " INNER JOIN DCIA010_借入金ワーク As K"
        wstr = wstr & " ON Z.借入番号=K.借入番号)"
        wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
        wstr = wstr & " ON K.銀行番号=G.銀行番号)"
        wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
        wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分)"
        wstr = wstr + " Left Join DAAA116_基準金利 As KK"
        wstr = wstr + " ON K.基準金利区分 = KK.基準金利区分"
    Else
        '金利SM
        wstr = wstr & " FROM (((DCDA010_借入残高推移表結果 As Z"
        wstr = wstr & " INNER JOIN DCIA010_借入金ワーク As K"
        wstr = wstr & " ON Z.借入番号=K.借入番号)"
        wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
        wstr = wstr & " ON K.銀行番号=G.銀行番号)"
        wstr = wstr & " LEFT JOIN DAAA115_金利シミュレーショングループ As KG"
        wstr = wstr & " ON K.金利グループ区分 = KG.金利グループ区分)"
        wstr = wstr + " Left Join DAAA116_基準金利 As KK"
        wstr = wstr + " ON K.基準金利区分 = KK.基準金利区分"
    End If
    
    'Where 条件
    'All 0 は表示しない
    wstr = wstr & " Where Z.前払利息増_" & w番号 & "<>0"
    wstr = wstr & " Or Z.前払利息減_" & w番号 & "<>0"
    wstr = wstr & " Or Z.前払利息_" & w番号 & "<>0"
    wstr = wstr & " Or Z.未払利息増_" & w番号 & "<>0"
    wstr = wstr & " Or Z.未払利息減_" & w番号 & "<>0"
    wstr = wstr & " Or Z.未払利息_" & w番号 & "<>0"
    
    If GStr <> "金利GR" Then
        wstr = wstr & " ORDER BY K.借入金種別区分,K.利息区分,K.銀行番号,K.借入番号"
    Else
        '金利SM
        wstr = wstr & " ORDER BY IIF(K.金利グループ区分<>'',K.金利グループ区分,'99999'),K.利息区分,K.銀行番号,K.借入番号"
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
    GLogStr = "推移表開始年度=" & GRpt.テキスト_01 & ","
    GLogStr = GLogStr & "推移表区分=" & GRpt.推移
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
    Me.Detail.Controls("I_融資残高") = Format(P8.FCDblRD(Me.Detail.Controls("I_融資残高")) / w分母, "#,##0")
    Me.Detail.Controls("I_前月利息残高") = Format(P8.FCDblRD(Me.Detail.Controls("I_前月利息残高")) / w分母, "#,##0")
    Me.Detail.Controls("I_利息増") = Format(P8.FCDblRD(Me.Detail.Controls("I_利息増")) / w分母, "#,##0")
    Me.Detail.Controls("I_利息減") = Format(P8.FCDblRD(Me.Detail.Controls("I_利息減")) / w分母, "#,##0")
    Me.Detail.Controls("I_利息残高") = Format(P8.FCDblRD(Me.Detail.Controls("I_利息残高")) / w分母, "#,##0")
'
End Sub

'------------------------------------------------
' GroupFooter1_BeforePrint
'------------------------------------------------
Private Sub GroupFooter1_BeforePrint()
'
    Me.GroupFooter1.Controls("G10_計名") = Me.GroupFooter1.Controls("G10_計名") & "　計"
'
    Me.GroupFooter1.Controls("G10_融資残高") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_融資残高")) / w分母, "#,##0")
    Me.GroupFooter1.Controls("G10_前月利息残高") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_前月利息残高")) / w分母, "#,##0")
    Me.GroupFooter1.Controls("G10_利息増") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_利息増")) / w分母, "#,##0")
    Me.GroupFooter1.Controls("G10_利息減") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_利息減")) / w分母, "#,##0")
    Me.GroupFooter1.Controls("G10_利息残高") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_利息残高")) / w分母, "#,##0")
'
End Sub

'------------------------------------------------
' GroupHeader2_BeforePrint
'------------------------------------------------
Private Sub GroupHeader2_BeforePrint()
'
    If Me.GroupHeader2.Controls("G20_NAME") = "利息先払" Then
        Me.GroupHeader2.Controls("G20_NAME") = "前払利息　計上"
        Me.GroupHeader2.Controls("L_利息増").Caption = "前払利息増"
        Me.GroupHeader2.Controls("L_利息減").Caption = "前払利息減"
        If GRpt.推移 = "月次" Then
            Me.GroupHeader2.Controls("L_利息前残").Caption = "前月前払利息残高"
            Me.GroupHeader2.Controls("L_利息残").Caption = "当月前払利息残高"
        Else
            Me.GroupHeader2.Controls("L_利息前残").Caption = "前期前払利息残高"
            Me.GroupHeader2.Controls("L_利息残").Caption = "当期前払利息残高"
        End If
    ElseIf Me.GroupHeader2.Controls("G20_NAME") = "利息後払" Then
        Me.GroupHeader2.Controls("G20_NAME") = "未払利息　計上"
        Me.GroupHeader2.Controls("L_利息増").Caption = "未払利息増"
        Me.GroupHeader2.Controls("L_利息減").Caption = "未払利息減"
        If GRpt.推移 = "月次" Then
            Me.GroupHeader2.Controls("L_利息前残").Caption = "前月未払利息残高"
            Me.GroupHeader2.Controls("L_利息残").Caption = "当月未払利息残高"
        Else
            Me.GroupHeader2.Controls("L_利息前残").Caption = "前期未払利息残高"
            Me.GroupHeader2.Controls("L_利息残").Caption = "当期未払利息残高"
        End If
    End If
'
End Sub

'------------------------------------------------
' GroupFooter2_BeforePrint
'------------------------------------------------
Private Sub GroupFooter2_BeforePrint()
'
    If Me.GroupFooter2.Controls("G20_計名") = "利息先払" Then
        Me.GroupFooter2.Controls("G20_計名") = "前払費用　計"
    ElseIf Me.GroupFooter2.Controls("G20_計名") = "利息後払" Then
        Me.GroupFooter2.Controls("G20_計名") = "未払費用　計"
    End If
'
    Me.GroupFooter2.Controls("G20_融資残高") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_融資残高")) / w分母, "#,##0")
    Me.GroupFooter2.Controls("G20_前月利息残高") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_前月利息残高")) / w分母, "#,##0")
    Me.GroupFooter2.Controls("G20_利息増") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_利息増")) / w分母, "#,##0")
    Me.GroupFooter2.Controls("G20_利息減") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_利息減")) / w分母, "#,##0")
    Me.GroupFooter2.Controls("G20_利息残高") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_利息残高")) / w分母, "#,##0")
'
End Sub

'------------------------------------------------
' GroupFooter3_BeforePrint
'------------------------------------------------
Private Sub GroupFooter3_BeforePrint()
'
    Me.GroupFooter3.Controls("G30_計名") = Me.GroupFooter3.Controls("G30_計名") & "　計"
'
    Me.GroupFooter3.Controls("G30_融資残高") = Format(P8.FCDblRD(Me.GroupFooter3.Controls("G30_融資残高")) / w分母, "#,##0")
    Me.GroupFooter3.Controls("G30_前月利息残高") = Format(P8.FCDblRD(Me.GroupFooter3.Controls("G30_前月利息残高")) / w分母, "#,##0")
    Me.GroupFooter3.Controls("G30_利息増") = Format(P8.FCDblRD(Me.GroupFooter3.Controls("G30_利息増")) / w分母, "#,##0")
    Me.GroupFooter3.Controls("G30_利息減") = Format(P8.FCDblRD(Me.GroupFooter3.Controls("G30_利息減")) / w分母, "#,##0")
    Me.GroupFooter3.Controls("G30_利息残高") = Format(P8.FCDblRD(Me.GroupFooter3.Controls("G30_利息残高")) / w分母, "#,##0")
'
End Sub

'------------------------------------------------
' ReportFooter_BeforePrint
'------------------------------------------------
Private Sub ReportFooter_BeforePrint()
'
    Me.ReportFooter.Controls("G90_融資残高") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_融資残高")) / w分母, "#,##0")
    Me.ReportFooter.Controls("G90_前月利息残高") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_前月利息残高")) / w分母, "#,##0")
    Me.ReportFooter.Controls("G90_利息増") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_利息増")) / w分母, "#,##0")
    Me.ReportFooter.Controls("G90_利息減") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_利息減")) / w分母, "#,##0")
    Me.ReportFooter.Controls("G90_利息残高") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_利息残高")) / w分母, "#,##0")
'
End Sub

'------------------------------------------------
' 金利GR_CHECK
'------------------------------------------------
Private Function 金利GR_CHECK() As Integer
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
'
    On Error GoTo 金利GR_CHECK_ERR
'
    wstr = "SELECT K.金利グループ区分"
    wstr = wstr & " FROM (DCDA010_借入残高推移表結果 AS Z"
    wstr = wstr & " INNER JOIN DCIA010_借入金ワーク AS K ON Z.借入番号 = K.借入番号)"
    wstr = wstr & " LEFT JOIN DAAA115_金利シミュレーショングループ AS KG"
    wstr = wstr & " ON K.金利グループ区分 = KG.金利グループ区分"
    wstr = wstr & " GROUP BY K.金利グループ区分"
    wstr = wstr & " Having K.金利グループ区分<>''"
    wstr = wstr & " ORDER BY K.金利グループ区分"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        金利GR_CHECK = wRs.RecordCount
    wRs.Close
    Set wRs = Nothing
'
    Exit Function
'
'----------< ERROR ROUTINE >--------------------------------------------------------
金利GR_CHECK_ERR:
    pERR_MES = pPROGRAM_ID + "/ 金利GR_CHECK() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

