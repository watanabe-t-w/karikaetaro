VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RDA040_借入金返済予定表 
   Caption         =   "借入金返済予定表"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10890
   Icon            =   "RDA040_借入金返済予定表.dsx":0000
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   _ExtentX        =   19209
   _ExtentY        =   15108
   SectionData     =   "RDA040_借入金返済予定表.dsx":0ECA
End
Attribute VB_Name = "RDA040_借入金返済予定表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "RDA040_借入金返済予定表"

Dim wstr As String, w円単位 As String
Dim w分母 As Integer
'Dim CNT_Line As Integer
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
    '----------------------------------------------------------------
    '                           ** 見出し **
    '----------------------------------------------------------------
    出力日 = Now
    企業名 = GCoName
    
    w分母 = 1
    L_単位 = "（円単位）"
    w円単位 = "円"
'
    L_帳票名.Caption = GRpt.テキスト_01 & "～" & GRpt.テキスト_02 & " 借入金返済予定表 "
'
    If GRpt.集計 = "年月日別" Then
            
        'GroupFooter1 銀行
        'GroupFooter2 年月日
        'GroupFooter3 年月
        GroupHeader1.DataField = "GrpFld_G"
        GroupHeader2.DataField = "GrpFld_D"
        GroupHeader3.DataField = "GrpFld_M"
        
        Me.GroupFooter2.Controls("G02_実際年月日").DataField = "I_実際年月日"
        
        Me.GroupFooter1.Controls("LG01_計名").Caption = "銀行計"
        Me.GroupFooter2.Controls("LG02_計名").Caption = "年月日計"
        Me.GroupFooter3.Controls("LG03_計名").Caption = "年月計"
        
        Me.GroupFooter3.Controls("G03_実際年月日").Visible = True
        
        Me.GroupFooter1.Controls("G01_銀行番号").Visible = True
        Me.GroupFooter2.Controls("G02_銀行番号").Visible = False
        Me.GroupFooter3.Controls("G03_銀行番号").Visible = False
        
        Me.GroupFooter1.Controls("G01_銀行名").Visible = True
        Me.GroupFooter2.Controls("G02_銀行名").Visible = False
        Me.GroupFooter3.Controls("G03_銀行名").Visible = False
        
        '銀行指定の場合はFooter2は表示しない
        GroupFooter2.Visible = False
        If GRpt.指定 = "" Then
            GroupFooter2.Visible = True
        End If
    ElseIf GRpt.集計 = "銀行別" Then
        'GroupFooter1 年月日
        'GroupFooter2 年月
        'GroupFooter3 銀行
        
        GroupHeader1.DataField = "GrpFld_D"
        GroupHeader2.DataField = "GrpFld_M"
        GroupHeader3.DataField = "GrpFld_G"

        Me.GroupFooter2.Controls("G02_実際年月日").DataField = "G03_実際年月日"
        
        Me.GroupFooter1.Controls("LG01_計名").Caption = "年月日計"
        Me.GroupFooter2.Controls("LG02_計名").Caption = "年月計"
        Me.GroupFooter3.Controls("LG03_計名").Caption = "銀行計"

        Me.GroupFooter3.Controls("G03_実際年月日").Visible = False
        
        Me.GroupFooter1.Controls("G01_銀行番号").Visible = True
        Me.GroupFooter2.Controls("G02_銀行番号").Visible = True
        Me.GroupFooter3.Controls("G03_銀行番号").Visible = True
        
        Me.GroupFooter1.Controls("G01_銀行名").Visible = True
        Me.GroupFooter2.Controls("G02_銀行名").Visible = True
        Me.GroupFooter3.Controls("G03_銀行名").Visible = True
    End If
'
    '印刷設定
    Me.Detail.Height = 0
    Me.PageHeader.Controls("L_PH借入番号").Visible = True
    Me.PageHeader.Controls("L_PH借入内容").Visible = True
    If GRpt.詳細表示 = 1 Then
        Me.Detail.Height = 220
        Me.PageHeader.Controls("L_PH借入番号").Visible = False
        Me.PageHeader.Controls("L_PH借入内容").Visible = False
    End If
'
    '----------------------------------------------------------------
    '          ** グループ ・ Where ・ OrderＢｙ **
    '----------------------------------------------------------------
    wWhere = ""
    
    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    FLG_TR = False
    
    '** 明細ファイル 削除 **
    wstr = ""
    wstr = wstr + "Delete * From DCDA020_借入金明細"
    GDb.Execute wstr
    
    '** 明細ファイル 作成 **
    '16/03/26 利子補給に伴う変更
    wstr = ""
    wstr = wstr & "Select K.*"
    wstr = wstr & " From DBDA010_借入金 As K"
    wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
    wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分"
    wstr = wstr & " Where K.取消フラグ=0"
    wstr = wstr & " And K.sm区分=0"
    wstr = wstr & " And S.利子補給金フラグ=0"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    Do Until wRs.EOF
        w借入データ = MBD010_借入データセット(wRs)
        If P8.FCDbl(wRs("手入力区分")) = "0" Then
        '手入力の場合は借入金データセットしない
            Call MBD010_借入金テーブル作成(w金融リストラ, w借入データ)
            Call MBD010_借入明細作成(w金融リストラ, w借入データ)        ' 07/02/21 V180
        Else
            ''手入力の場合は下記の明細TR作成を仕様
            'FLG_TR = True
        
            Call MBD010_借入金入力明細Read(w借入データ)
            If w借入データ.社債フラグ = 1 Then
                Call MDA020_借入金入力社債明細作成(w借入データ)
            End If
            Call MBD010_借入明細作成_入力登録(w借入データ)
        
        End If
        wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing

    'If FLG_TR = True Then
    '    Call MBD010_借入明細作成_明細TR(GRpt.コンボ_01, "DBDA010_借入金明細TR")
    'End If
'
    '----------------------------------------------------------------
    '                          ** 合計 計算 **
    '----------------------------------------------------------------
    
    '----------------------------------------------------------------
    '                       ** レコード　ソース **
    '----------------------------------------------------------------
    wstr = ""
    wstr = wstr & "Select"
    wstr = wstr & " format(M.実際年月日,'" & Gfmt年月日 & "') As I_実際年月日,"
    wstr = wstr & " G.銀行番号 As I_銀行番号,"
    wstr = wstr & " G.銀行番号 As GrpFld_G,"
    wstr = wstr & " M.実際年月日 As GrpFld_D,"
    wstr = wstr & " Format(M.実際年月日,'yyyy/mm') As GrpFld_M,"
    wstr = wstr & " Format(M.実際年月日,'yyyy/mm') As G03_実際年月日,"
    wstr = wstr & " G.銀行名 As I_銀行名,"
    wstr = wstr & " M.借入番号 As I_借入番号,"
    wstr = wstr & " K.借入内容 As I_借入内容,"
    wstr = wstr & " M.元金額 As I_元金額,"
    wstr = wstr & " M.返済金額 As I_返済金額,"
    wstr = wstr & " M.利息額 As I_利息額,"
    wstr = wstr & " M.初期手数料+M.元金手数料+M.利息手数料 As I_手数料,"
    wstr = wstr & " M.保証料 As I_保証料"
    
    wstr = wstr & " FROM ((DCDA020_借入金明細 AS M"
    wstr = wstr & " INNER JOIN DBDA010_借入金 AS K"
    wstr = wstr & " ON M.借入番号 = K.借入番号)"
    wstr = wstr & " INNER JOIN DAAA040_銀行マスタ AS G"
    wstr = wstr & " ON K.銀行番号 = G.銀行番号)"
    wstr = wstr & " INNER JOIN DAAA116_借入金種別 AS S"
    wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分"
    
    GVar1 = C年月日.平成To西暦("年月日", GRpt.テキスト_01)
    GVar2 = C年月日.平成To西暦("年月日", GRpt.テキスト_02)
    
'    wstr = wstr & " Where M.返済金額<>0" '初回返済で利息前払い分除外
'    wstr = wstr & " And M.融資残高>=0" '内入れ金額のマイナス分除外
    
'    If GRpt.テキスト_01 <> "" And GRpt.テキスト_02 = "" Then
'        wstr = wstr & " And Format(M.実際年月日,'yyyy/mm/dd')>='" & Format(GVar1, "yyyy/mm/dd") & "'"
'    ElseIf GRpt.テキスト_01 = "" And GRpt.テキスト_02 <> "" Then
'        wstr = wstr & " And Format(M.実際年月日,'yyyy/mm/dd')<='" & Format(GVar2, "yyyy/mm/dd") & "'"
'    ElseIf GRpt.テキスト_01 <> "" And GRpt.テキスト_02 <> "" Then
'        wstr = wstr & " And Format(M.実際年月日,'yyyy/mm/dd')>='" & Format(GVar1, "yyyy/mm/dd") & "'"
'        wstr = wstr & " And Format(M.実際年月日,'yyyy/mm/dd')<='" & Format(GVar2, "yyyy/mm/dd") & "'"
'    End If
'
'    If GRpt.指定 <> "" Then
'        wstr = wstr & " And G.銀行名='" & GRpt.指定 & "'"
'    End If
        
    wstr = wstr & " WHERE (M.返済金額<>0"
    wstr = wstr & " AND M.融資残高>=0"
    If GRpt.指定 <> "" Then
        wstr = wstr & " And G.銀行名='" & GRpt.指定 & "'"
    End If
    If GRpt.テキスト_01 <> "" And GRpt.テキスト_02 = "" Then
        wstr = wstr & " And Format(M.実際年月日,'yyyy/mm/dd')>='" & Format(GVar1, "yyyy/mm/dd") & "'"
    ElseIf GRpt.テキスト_01 = "" And GRpt.テキスト_02 <> "" Then
        wstr = wstr & " And Format(M.実際年月日,'yyyy/mm/dd')<='" & Format(GVar2, "yyyy/mm/dd") & "'"
    ElseIf GRpt.テキスト_01 <> "" And GRpt.テキスト_02 <> "" Then
        wstr = wstr & " And Format(M.実際年月日,'yyyy/mm/dd')>='" & Format(GVar1, "yyyy/mm/dd") & "'"
        wstr = wstr & " And Format(M.実際年月日,'yyyy/mm/dd')<='" & Format(GVar2, "yyyy/mm/dd") & "'"
    End If
    wstr = wstr & ")"
    
    '社債の手数料
    wstr = wstr & " OR ((M.初期手数料+M.元金手数料+M.利息手数料<>0)"
    If GRpt.指定 <> "" Then
        wstr = wstr & " And G.銀行名='" & GRpt.指定 & "'"
    End If
    If GRpt.テキスト_01 <> "" And GRpt.テキスト_02 = "" Then
        wstr = wstr & " And Format(M.実際年月日,'yyyy/mm/dd')>='" & Format(GVar1, "yyyy/mm/dd") & "'"
    ElseIf GRpt.テキスト_01 = "" And GRpt.テキスト_02 <> "" Then
        wstr = wstr & " And Format(M.実際年月日,'yyyy/mm/dd')<='" & Format(GVar2, "yyyy/mm/dd") & "'"
    ElseIf GRpt.テキスト_01 <> "" And GRpt.テキスト_02 <> "" Then
        wstr = wstr & " And Format(M.実際年月日,'yyyy/mm/dd')>='" & Format(GVar1, "yyyy/mm/dd") & "'"
        wstr = wstr & " And Format(M.実際年月日,'yyyy/mm/dd')<='" & Format(GVar2, "yyyy/mm/dd") & "'"
    End If
    wstr = wstr & " AND S.社債フラグ=1)"
    
    '社債の保証料
    wstr = wstr & " OR (M.保証料<>0"
    If GRpt.指定 <> "" Then
        wstr = wstr & " And G.銀行名='" & GRpt.指定 & "'"
    End If
    If GRpt.テキスト_01 <> "" And GRpt.テキスト_02 = "" Then
        wstr = wstr & " And Format(M.実際年月日,'yyyy/mm/dd')>='" & Format(GVar1, "yyyy/mm/dd") & "'"
    ElseIf GRpt.テキスト_01 = "" And GRpt.テキスト_02 <> "" Then
        wstr = wstr & " And Format(M.実際年月日,'yyyy/mm/dd')<='" & Format(GVar2, "yyyy/mm/dd") & "'"
    ElseIf GRpt.テキスト_01 <> "" And GRpt.テキスト_02 <> "" Then
        wstr = wstr & " And Format(M.実際年月日,'yyyy/mm/dd')>='" & Format(GVar1, "yyyy/mm/dd") & "'"
        wstr = wstr & " And Format(M.実際年月日,'yyyy/mm/dd')<='" & Format(GVar2, "yyyy/mm/dd") & "'"
    End If
    wstr = wstr & " AND S.社債フラグ=1)"
    
    If GRpt.集計 = "年月日別" Then
        wstr = wstr & " ORDER BY M.実際年月日,K.銀行番号,M.借入番号,M.据置X回目"
    ElseIf GRpt.集計 = "銀行別" Then
        wstr = wstr & " ORDER BY K.銀行番号,M.実際年月日,M.借入番号,M.据置X回目"
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
