VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RDH040_支払利息推移表 
   Caption         =   "支払利息推移表"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12480
   Icon            =   "RDH040_支払利息推移表.dsx":0000
   StartUpPosition =   3  'Windows の既定値
   WindowState     =   2  '最大化
   _ExtentX        =   22013
   _ExtentY        =   12806
   SectionData     =   "RDH040_支払利息推移表.dsx":0ECA
End
Attribute VB_Name = "RDH040_支払利息推移表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'杉村倉庫仕様
Private Const pPROGRAM_ID As String = "RDH040_支払利息推移表"
'
Dim wML As Integer, wML2 As Integer
Dim w番号 As String, wsTbl As String, wsTbl2 As String
Dim w分母 As Integer
Dim w推移表タイトル As MAA910_推移表タイトル
'
'------------------------------------------------
' ActiveReport_ReportStart
'------------------------------------------------
Private Sub ActiveReport_ReportStart()
'
    Dim w開始年月日 As Date, wdate As Date
    Dim w推移表区分 As String
    Dim j As Integer, wIndex As Integer
    Dim wsS As String
    Dim ws01 As String
    Dim wstr As String
    Dim wOrder As String
    Dim FLG_Order As Boolean
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
    
    wML = 12
'
    '----------------------------------------------------------------
    '                           ** 見出し **
    '----------------------------------------------------------------
    出力日 = Now
    企業名 = GCoName
    
    L_帳票名.Caption = GRpt.テキスト_01 & "年度" & Left(GRpt.テキスト_02, 5) & " " & GRpt.帳票名
    
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
    wsS = wsS & "帳票指示:"
'
    '----------------------------------------------------------------
    '          ** グループ ・ Where ・ OrderＢｙ **
    '----------------------------------------------------------------
    'グループセット
    'グループセット
    If GRpt.S_種別 = "分類1" Then
        GroupHeader1.DataField = "GrpFld_KShubetu"
    ElseIf GRpt.S_種別 = "分類2" Then
        GroupHeader2.DataField = "GrpFld_KShubetu"
    ElseIf GRpt.S_種別 = "分類3" Then
        GroupHeader3.DataField = "GrpFld_KShubetu"
    End If
    
    If GRpt.S_銀行 = "分類1" Then
        GroupHeader1.DataField = "GrpFld_Ginko"
    ElseIf GRpt.S_銀行 = "分類2" Then
        GroupHeader2.DataField = "GrpFld_Ginko"
    ElseIf GRpt.S_銀行 = "分類3" Then
        GroupHeader3.DataField = "GrpFld_Ginko"
    End If
    
    If GRpt.S_利息 = "分類1" Then
        GroupHeader1.DataField = "GrpFld_Risoku"
    ElseIf GRpt.S_利息 = "分類2" Then
        GroupHeader2.DataField = "GrpFld_Risoku"
    ElseIf GRpt.S_利息 = "分類3" Then
        GroupHeader3.DataField = "GrpFld_Risoku"
    End If
    
    '計名セット、Shapeカラー
    If GroupHeader1.DataField = "GrpFld_KShubetu" Then
        wsS = wsS & "分類1:借入金種別 "
        Me.GroupFooter1.Controls("G10_計名").DataField = "I_借入金種別名"
        Me.GroupFooter1.Controls("Shape_1").BackColor = &HFFFFC0
    ElseIf GroupHeader1.DataField = "GrpFld_Ginko" Then
        wsS = wsS & "分類1:銀行 "
        Me.GroupFooter1.Controls("G10_計名").DataField = "I_銀行名"
        Me.GroupFooter1.Controls("Shape_1").BackColor = &HC0FFFF
    ElseIf GroupHeader1.DataField = "GrpFld_Risoku" Then
        wsS = wsS & "分類1:利息区分 "
        Me.GroupFooter1.Controls("G10_計名").DataField = "I_利息区分"
        Me.GroupFooter1.Controls("Shape_1").BackColor = &HE0E0E0
    End If
    
    If GroupHeader2.DataField = "GrpFld_KShubetu" Then
        wsS = wsS & "分類2:借入金種別 "
        Me.GroupFooter2.Controls("G20_計名").DataField = "I_借入金種別名"
        Me.GroupFooter2.Controls("Shape_2").BackColor = &HFFFFC0
    ElseIf GroupHeader2.DataField = "GrpFld_Ginko" Then
        wsS = wsS & "分類2:銀行 "
        Me.GroupFooter2.Controls("G20_計名").DataField = "I_銀行名"
        Me.GroupFooter2.Controls("Shape_2").BackColor = &HC0FFFF
    ElseIf GroupHeader2.DataField = "GrpFld_Risoku" Then
        wsS = wsS & "分類2:利息区分 "
        Me.GroupFooter2.Controls("G20_計名").DataField = "I_利息区分"
        Me.GroupFooter2.Controls("Shape_2").BackColor = &HE0E0E0
    End If
    
    If GroupHeader3.DataField = "GrpFld_KShubetu" Then
        wsS = wsS & "分類3:借入金種別 "
        Me.GroupFooter3.Controls("G30_計名").DataField = "I_借入金種別名"
        Me.GroupFooter3.Controls("Shape_3").BackColor = &HFFFFC0
    ElseIf GroupHeader3.DataField = "GrpFld_Ginko" Then
        wsS = wsS & "分類3:銀行 "
        Me.GroupFooter3.Controls("G30_計名").DataField = "I_銀行名"
        Me.GroupFooter3.Controls("Shape_3").BackColor = &HC0FFFF
    ElseIf GroupHeader3.DataField = "GrpFld_Risoku" Then
        wsS = wsS & "分類3:利息区分 "
        Me.GroupFooter3.Controls("G30_計名").DataField = "I_利息区分"
        Me.GroupFooter3.Controls("Shape_3").BackColor = &HE0E0E0
    End If
'
    GroupFooter1.Visible = True
    GroupFooter2.Visible = True
    GroupFooter3.Visible = True
    If GroupHeader1.DataField = "" Then
        GroupFooter1.Visible = False
    End If
    If GroupHeader2.DataField = "" Then
        GroupFooter2.Visible = False
    End If
    If GroupHeader3.DataField = "" Then
        GroupFooter3.Visible = False
    End If
    
    '改ページ
    Me.GroupFooter1.NewPage = GRpt.NewPage1
    Me.GroupFooter2.NewPage = GRpt.NewPage2
    Me.GroupFooter3.NewPage = GRpt.NewPage3
'
    '帳票指示
    Me.PageHeader.Controls("L_帳票指示").Caption = wsS
'
    '印刷設定
    If GRpt.詳細表示 = 1 Then
        Me.Detail.Height = 218
        Me.PageHeader.Controls("L_番号") = "借入番号"
        Me.PageHeader.Controls("L_利率").Visible = True
        Me.PageHeader.Controls("Line_H1").Visible = True
    Else
        Me.Detail.Height = 0
        Me.PageHeader.Controls("L_番号") = "銀行名"
        Me.PageHeader.Controls("L_利率").Visible = False
        Me.PageHeader.Controls("Line_H1").Visible = False
    End If
'
    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    Call MBD020_借入金ワークテーブル作成(wsTbl) 'データ絞り込み
    Call MRB010_標準入力借入残高表("DCIA010_借入金ワーク")   '16/03/26利子補給に伴う変更
    'Call MRB010_標準入力借入残高表固定日数("DCIA010_借入金ワーク")
    Call MRB010_手入力借入残高表("DCIA010_借入金ワーク")
'
    w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "西暦")
    w推移表タイトル = MUA010_推移表ファイル作成("", "", w開始年月日, GRpt.推移, wML)

    '杉村倉庫仕様
    wML = 4
    wML2 = CInt(Mid(GRpt.テキスト_02, 2, 1))
    
    Call RDB020_コントロールセット

    'CSVファイル パラメータ:GInt1
    GInt1 = wML2
'
    '----------------------------------------------------------------
    '                          ** 合計 計算 **
    '----------------------------------------------------------------
'
    '----------------------------------------------------------------
    '                       ** レコード　ソース **
    '----------------------------------------------------------------
    '** レコード　ソース **
    wstr = "Select "
    wstr = wstr & "K.借入番号 As I_借入番号,"
    
    'セクションGR
    wstr = wstr & "K.利息区分 As GrpFld_Risoku,"
    wstr = wstr & "K.銀行番号 As GrpFld_Ginko,"
    wstr = wstr & "K.借入金種別区分 As GrpFld_KShubetu,"
    
    wstr = wstr & "IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As I_利息区分,"
    wstr = wstr & "G.銀行名 As I_銀行名,"
    wstr = wstr & "S.借入金種別名 As I_借入金種別名,"
    
    w番号 = Right("00" + CStr(wML2), 2)
    wstr = wstr & "利率_" + w番号 + " As I_利率,"
    
    wstr = wstr & "K.融資金額 As I_融資金額,"
    
    '杉村倉庫仕様
    For j = 1 To wML2
        w番号 = Right("00" + CStr(j), 2)
        wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & ",Z.未払利息増_" & w番号 & ") As I_" + w番号 + "1,"
    Next
    
    '合計
    w番号 = "01"
    wstr = wstr & "(IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & ",Z.未払利息増_" & w番号 & ")"
    For j = 2 To wML2
        w番号 = Right("00" + CStr(j), 2)
        wstr = wstr & " + IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & ",Z.未払利息増_" & w番号 & ")"
    Next
    wstr = wstr & ") As I_001"
    
    'wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減合計,Z.未払利息増合計) As I_001"
    
    wstr = wstr & " FROM ((((DCDA010_借入残高推移表結果 As Z"
    wstr = wstr & " INNER JOIN DCDA010_借入残高推移表結果２ As Z2"
    wstr = wstr & " ON Z.借入番号=Z2.借入番号)"
    wstr = wstr & " INNER JOIN DCIA010_借入金ワーク As K"
    wstr = wstr & " ON Z.借入番号=K.借入番号)"
    wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr & " ON K.銀行番号=G.銀行番号)"
    wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
    wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分)"
    
    'Where 条件
    'All 0 は表示しない
    wstr = wstr & " Where ("
    '融資
    wstr = wstr & " Z.融資_01<>0"
    For j = 2 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.融資_" & ws01 & "<>0"
    Next j
    '元金
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.元金_" & ws01 & "<>0"
    Next j
    '利息
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.利息_" & ws01 & "<>0"
    Next j
    '返済
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.返済_" & ws01 & "<>0"
    Next j
    '解約
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.解約_" & ws01 & "<>0"
    Next j
    '保証
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.保証_" & ws01 & "<>0"
    Next j
    '残高
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.残高_" & ws01 & "<>0"
    Next j
    '前払利息
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.前払利息_" & ws01 & "<>0"
    Next j
    '前払利息増
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.前払利息増_" & ws01 & "<>0"
    Next j
    '前払利息減
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.前払利息減_" & ws01 & "<>0"
    Next j
    '未払利息
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.未払利息_" & ws01 & "<>0"
    Next j
    '未払利息増
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.未払利息増_" & ws01 & "<>0"
    Next j
    '未払利息減
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.未払利息減_" & ws01 & "<>0"
    Next j
    '損益利息額
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z2.損益利息額_" & ws01 & "<>0"
    Next j
    wstr = wstr & " )"
        
    wOrder = "": FLG_Order = False
    For j = 1 To 4
        If (GRpt.S_金融 = "分類" & CStr(j) Or GRpt.S_銀行 = "分類" & CStr(j)) _
        And FLG_Order = False Then
            wOrder = wOrder & "K.銀行番号,"
            FLG_Order = True
        ElseIf GRpt.S_種別 = "分類" & CStr(j) Then
            wOrder = wOrder & "K.借入金種別区分,"
        ElseIf GRpt.S_利息 = "分類" & CStr(j) Then
            wOrder = wOrder & "K.利息区分,"
        End If
    Next j
    wOrder = " Order by " & wOrder & "K.借入番号"
    wstr = wstr & wOrder
    
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
' RDB020_コントロールセット
'------------------------------------------------
Private Sub RDB020_コントロールセット()
    Dim j As Integer
'
    For j = 1 To wML
        w番号 = Right("00" + CStr(j), 2)
        Me.PageHeader.Controls("Lbl_" + w番号 + "番目年月") = w推移表タイトル.X番目年月(j)
    Next
'
End Sub

'------------------------------------------------
' Detail_BeforePrint
'------------------------------------------------
Private Sub Detail_BeforePrint()
'
    Dim j As Integer
    Dim wstr As String
'
    Me.Detail.Controls("I_001") = Format(P8.FCDblRD(Me.Detail.Controls("I_001")) / w分母, "#,##0")
    Me.Detail.Controls("I_融資金額") = Format(P8.FCDblRD(Me.Detail.Controls("I_融資金額")) / w分母, "#,##0")
    Call MXA030_ReportColor(Me.Detail.Controls("I_001"))
    Call MXA030_ReportColor(Me.Detail.Controls("I_融資金額"))
    
    For j = 0 To wML
        wstr = Right("00" + CStr(j), 2) + "1"
        
        '杉村倉庫仕様
        If j <= wML2 Then
            Me.Detail.Controls("I_" + wstr) = Format(P8.FCDblRD(Me.Detail.Controls("I_" + wstr)) / w分母, "#,##0")
        Else
            Me.Detail.Controls("I_" + wstr) = ""
        End If
        Call MXA030_ReportColor(Me.Detail.Controls("I_" + wstr))
    Next j
'
End Sub

'------------------------------------------------
' GroupFooter1_BeforePrint
'------------------------------------------------
Private Sub GroupFooter1_BeforePrint()
'
    Dim j As Integer
    Dim wstr As String
'
    Me.GroupFooter1.Controls("G10_001") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_001")) / w分母, "#,##0")
    Me.GroupFooter1.Controls("G10_融資金額") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_融資金額")) / w分母, "#,##0")
    Call MXA030_ReportColor(Me.GroupFooter1.Controls("G10_001"))
    Call MXA030_ReportColor(Me.GroupFooter1.Controls("G10_融資金額"))
    
    For j = 0 To wML
        wstr = Right("00" + CStr(j), 2) + "1"
        
        '杉村倉庫仕様
        If j <= wML2 Then
            Me.GroupFooter1.Controls("G10_" + wstr) = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_" + wstr)) / w分母, "#,##0")
        Else
            Me.GroupFooter1.Controls("G10_" + wstr) = ""
        End If
        Call MXA030_ReportColor(Me.GroupFooter1.Controls("G10_" + wstr))
    Next j
'
End Sub

'------------------------------------------------
' GroupFooter2_BeforePrint
'------------------------------------------------
Private Sub GroupFooter2_BeforePrint()
'
    Dim j As Integer
    Dim wstr As String
'
    Me.GroupFooter2.Controls("G20_001") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_001")) / w分母, "#,##0")
    Me.GroupFooter2.Controls("G20_融資金額") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_融資金額")) / w分母, "#,##0")
    Call MXA030_ReportColor(Me.GroupFooter2.Controls("G20_001"))
    Call MXA030_ReportColor(Me.GroupFooter2.Controls("G20_融資金額"))
    
    For j = 0 To wML
        wstr = Right("00" + CStr(j), 2) + "1"
        
        '杉村倉庫仕様
        If j <= wML2 Then
            Me.GroupFooter2.Controls("G20_" + wstr) = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_" + wstr)) / w分母, "#,##0")
        Else
            Me.GroupFooter2.Controls("G20_" + wstr) = ""
        End If
        Call MXA030_ReportColor(Me.GroupFooter2.Controls("G20_" + wstr))
    Next j
'
End Sub

'------------------------------------------------
' GroupFooter3_BeforePrint
'------------------------------------------------
Private Sub GroupFooter3_BeforePrint()
'
    Dim j As Integer
    Dim wstr As String
'
    Me.GroupFooter3.Controls("G30_001") = Format(P8.FCDblRD(Me.GroupFooter3.Controls("G30_001")) / w分母, "#,##0")
    Me.GroupFooter3.Controls("G30_融資金額") = Format(P8.FCDblRD(Me.GroupFooter3.Controls("G30_融資金額")) / w分母, "#,##0")
    Call MXA030_ReportColor(Me.GroupFooter3.Controls("G30_001"))
    Call MXA030_ReportColor(Me.GroupFooter3.Controls("G30_融資金額"))
    
    For j = 0 To wML
        wstr = Right("00" + CStr(j), 2) + "1"
        
        '杉村倉庫仕様
        If j <= wML2 Then
            Me.GroupFooter3.Controls("G30_" + wstr) = Format(P8.FCDblRD(Me.GroupFooter3.Controls("G30_" + wstr)) / w分母, "#,##0")
        Else
            Me.GroupFooter3.Controls("G30_" + wstr) = ""
        End If
        Call MXA030_ReportColor(Me.GroupFooter3.Controls("G30_" + wstr))
    Next j
'
End Sub

'------------------------------------------------
' ReportFooter_BeforePrint
'------------------------------------------------
Private Sub ReportFooter_BeforePrint()
'
    Dim j As Integer
    Dim wstr As String
'
    Me.ReportFooter.Controls("G90_001") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_001")) / w分母, "#,##0")
    Me.ReportFooter.Controls("G90_融資金額") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_融資金額")) / w分母, "#,##0")
    Call MXA030_ReportColor(Me.ReportFooter.Controls("G90_001"))
    Call MXA030_ReportColor(Me.ReportFooter.Controls("G90_融資金額"))
    
    For j = 0 To wML
        wstr = Right("00" + CStr(j), 2) + "1"
        
        '杉村倉庫仕様
        If j <= wML2 Then
            Me.ReportFooter.Controls("G90_" + wstr) = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_" + wstr)) / w分母, "#,##0")
        Else
            Me.ReportFooter.Controls("G90_" + wstr) = ""
        End If
        Call MXA030_ReportColor(Me.ReportFooter.Controls("G90_" + wstr))
    Next j
'
End Sub
