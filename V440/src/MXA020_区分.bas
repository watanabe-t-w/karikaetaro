Attribute VB_Name = "MXA020_区分"
Option Explicit
'
Private Const pPROGRAM_ID As String = "MXA020_区分"
'
'------------------------------------------------
' XMXA020_区分
'------------------------------------------------
Public Function XMXA020_区分(p区分名 As String, pコード) As String
'
    On Error GoTo XMXA020_区分_ERR
'
    ' =========================================
    '                  資産区分
    ' =========================================
    If p区分名 = "資産区分" Then
        Select Case pコード
            Case "有形資産": XMXA020_区分 = "1"
            Case "無形資産": XMXA020_区分 = "2"
            Case "損金設備": XMXA020_区分 = "3"
            Case "建物": XMXA020_区分 = "4"
            Case "土地": XMXA020_区分 = "5"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               借入金管理区分
    ' =========================================
    If p区分名 = "借入金管理区分" Then
        Select Case pコード
            Case "管理用": XMXA020_区分 = "0"
            Case "決算用": XMXA020_区分 = "1"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               決算サイクル
    ' =========================================
    If p区分名 = "決算サイクル" Then
        Select Case pコード
            Case "月次": XMXA020_区分 = "1"
            Case "四半期": XMXA020_区分 = "3"
            Case "半期": XMXA020_区分 = "6"
            Case "年次": XMXA020_区分 = "12"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               消費税納税条件
    ' =========================================
    If p区分名 = "消費税納税条件" Then
        Select Case pコード
            Case "1ヶ月毎": XMXA020_区分 = "1"
            Case "3ヶ月毎": XMXA020_区分 = "3"
            Case "6ヶ月毎": XMXA020_区分 = "6"
            Case "12ヶ月毎": XMXA020_区分 = "12"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               予算設定区分
    ' =========================================
    If p区分名 = "予算設定区分" Then
        Select Case pコード
            Case "消費税抜": XMXA020_区分 = "1"
            Case "消費税込": XMXA020_区分 = "2"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               決算書参照区分
    ' =========================================
    If p区分名 = "決算書参照区分" Then
        Select Case pコード
            Case "経費、資金状況": XMXA020_区分 = "1"
            Case "全ての項目対象": XMXA020_区分 = "2"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '             減価償却費計上
    ' =========================================
    If p区分名 = "減価償却費計上" Then
        Select Case pコード
            Case "毎月計上": XMXA020_区分 = "1"
            Case "期末計上": XMXA020_区分 = "2"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               業種区分
    ' =========================================
    If p区分名 = "業種区分" Then
        Select Case pコード
            Case "非製造業": XMXA020_区分 = "1"
            Case "製造業": XMXA020_区分 = "2"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '                  償却区分
    ' =========================================
    If p区分名 = "償却区分" Then
        Select Case pコード
            Case "定額法": XMXA020_区分 = "1"
            Case "定率法": XMXA020_区分 = "2"
            Case "均等償却": XMXA020_区分 = "3"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '                  減価償却費区分
    ' =========================================
    If p区分名 = "減価償却費区分" Then
        Select Case pコード
            Case "一般管理費": XMXA020_区分 = "1"
            Case "製造原価": XMXA020_区分 = "2"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '                  課税区分
    ' =========================================
    If p区分名 = "課税区分" Then
        If Len(pコード) > 1 Then
            Select Case pコード
                Case "不課税": XMXA020_区分 = "1"
                Case "課税": XMXA020_区分 = "2"
                Case "非課税": XMXA020_区分 = "3"
            End Select
        Else
            Select Case pコード
                Case "1": XMXA020_区分 = "不課税"
                Case "2": XMXA020_区分 = "課税"
                Case "3": XMXA020_区分 = "非課税"
            End Select
        End If
        Exit Function
    End If
'
    ' =========================================
    '                  現金取引
    ' =========================================
    If p区分名 = "現金取引" Then
        Select Case pコード
            Case "現金のみ": XMXA020_区分 = "1"
            Case "掛売有り": XMXA020_区分 = "2"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '                  営業日
    ' =========================================
    If p区分名 = "営業日" Then
        Select Case pコード
            Case "翌営業日": XMXA020_区分 = "0"
            Case "前営業日": XMXA020_区分 = "1"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '                  利息区分
    ' =========================================
    If p区分名 = "利息区分" Then
        If Len(pコード) > 1 Then
            Select Case pコード
                Case "利息先払": XMXA020_区分 = "1"
                Case "利息後払": XMXA020_区分 = "2"
            End Select
        Else
            Select Case pコード
                Case "1": XMXA020_区分 = "利息先払"
                Case "2": XMXA020_区分 = "利息後払"
            End Select
        End If
        Exit Function
    End If
'
    ' =========================================
    '               利息計算日数
    ' =========================================
    If p区分名 = "利息日数" Then
        Select Case pコード
            Case "営業日数": XMXA020_区分 = "0"
            Case "固定日数": XMXA020_区分 = "1"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               利息支払
    ' =========================================
    If p区分名 = "利息支払" Then
        Select Case pコード
            Case "毎月": XMXA020_区分 = "0"
            Case "一括": XMXA020_区分 = "1"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               利息控除
    ' =========================================
    If p区分名 = "利息控除" Then
        Select Case pコード
            Case "控除無し":                XMXA020_区分 = "0"
            Case "実行日控除":              XMXA020_区分 = "1"
            Case "最終返済日控除":          XMXA020_区分 = "2"
            Case "実行日及び最終返済日控除": XMXA020_区分 = "3"
            Case "中間利払最終日控除":      XMXA020_区分 = "4"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               金利計算
    ' =========================================
    If p区分名 = "金利計算" Then
        Select Case pコード
            Case "365日": XMXA020_区分 = "0"
            Case "360日": XMXA020_区分 = "1"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               設備区分
    ' =========================================
    If p区分名 = "設備区分" Then
        If Len(pコード) > 1 Then
            Select Case pコード
                Case "運転資金": XMXA020_区分 = "0"
                Case "設備": XMXA020_区分 = "1"
            End Select
        Else
            Select Case pコード
                Case "0": XMXA020_区分 = "運転資金"
                Case "1": XMXA020_区分 = "設備"
            End Select
        End If
        Exit Function
    End If
'
    ' =========================================
    '               有担フラグ
    ' =========================================
    If p区分名 = "有担フラグ" Then
        If Len(pコード) > 1 Then
            Select Case pコード
                Case "無担保": XMXA020_区分 = "0"
                Case "有担保": XMXA020_区分 = "1"
            End Select
        Else
            Select Case pコード
                Case "0": XMXA020_区分 = "無担保"
                Case "1": XMXA020_区分 = "有担保"
            End Select
        End If
        Exit Function
    End If
'
    ' =========================================
    '               金利種別
    ' =========================================
    If p区分名 = "金利種別" Then
        If Len(pコード) > 1 Then
            Select Case pコード
                Case "変動金利": XMXA020_区分 = "0"
                Case "固定金利": XMXA020_区分 = "1"
            End Select
        Else
            Select Case pコード
                Case "0": XMXA020_区分 = "変動金利"
                Case "1": XMXA020_区分 = "固定金利"
            End Select
        End If
        Exit Function
    End If
'
    ' =========================================
    '               長短区分
    ' =========================================
    If p区分名 = "長短区分" Then
        If Len(pコード) > 1 Then
            If Len(pコード) > 2 Then
                Select Case pコード
                    Case "短期借入金": XMXA020_区分 = "0"
                    Case "長期借入金": XMXA020_区分 = "1"
                End Select
            ElseIf Len(pコード) = 2 Then
                Select Case pコード
                    Case "短期": XMXA020_区分 = "0"
                    Case "長期": XMXA020_区分 = "1"
                End Select
            End If
        Else
            Select Case pコード
                Case "0": XMXA020_区分 = "短期借入金"
                Case "1": XMXA020_区分 = "長期借入金"
            End Select
        End If
        Exit Function
    End If
'
    ' =========================================
    '               借換区分
    ' =========================================
    If p区分名 = "借換区分" Then
        Select Case pコード
            Case "無": XMXA020_区分 = "0"
            Case "借換": XMXA020_区分 = "1"
            Case "完済": XMXA020_区分 = "2"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               資金調達区分
    ' =========================================
    If p区分名 = "資金調達区分" Then
        Select Case pコード
            Case "本部":        XMXA020_区分 = "0"
            Case "支店":        XMXA020_区分 = "1"
            Case "単独企業":    XMXA020_区分 = "2"
            Case "連結本部":    XMXA020_区分 = "5"
            Case "連結子会社":  XMXA020_区分 = "6"
            Case "連結親会社":  XMXA020_区分 = "7"
            Case "全社":        XMXA020_区分 = "9"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '              シミュレーション
    ' =========================================
    If p区分名 = "シミュレーション" Then
        Select Case pコード
            Case "通常":            XMXA020_区分 = "0"
            Case "シミュレーション": XMXA020_区分 = "1"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '              金剛石科目番号
    ' =========================================
    If p区分名 = "金剛石科目番号" Then
        Select Case pコード
            Case "売上額":      XMXA020_区分 = "10100"
            Case "粗利益":      XMXA020_区分 = "10200"
            Case "給与総額":    XMXA020_区分 = "11100"
            Case "賞与額":      XMXA020_区分 = "11150"
            Case "固定経費":    XMXA020_区分 = "11200"
            Case "変動経費1":   XMXA020_区分 = "11250"
            Case "変動経費2":   XMXA020_区分 = "11300"
            Case "変動経費3":   XMXA020_区分 = "11320"
            Case "その他経費1": XMXA020_区分 = "11500"
            Case "保険積立":    XMXA020_区分 = "11550"
            Case "営業外収益":  XMXA020_区分 = "13100"
            Case "営業外費用":  XMXA020_区分 = "13200"
            Case "減価償却":    XMXA020_区分 = "11350"
            Case "支払利息":    XMXA020_区分 = "13300"
'            Case "売掛金":      XMXA020_区分 = "21000"
            Case "総債権残高":  XMXA020_区分 = "20810"
            Case "総債務残高":  XMXA020_区分 = "20830"
            Case "その他債権残高":  XMXA020_区分 = "20910"
            Case "その他債務残高":  XMXA020_区分 = "20920"
            Case "決算時手持資金": XMXA020_区分 = "50100"
            Case "決算時保険等": XMXA020_区分 = "50200"
            Case "決算時不動産証券": XMXA020_区分 = "50300"
            Case "決算時含資産": XMXA020_区分 = "50400"
            Case "決算時在庫": XMXA020_区分 = "50500"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '              決算書区分
    ' =========================================
    If p区分名 = "決算書区分" Then
        Select Case pコード
            Case "貸借対照表":  XMXA020_区分 = "1"
            Case "損益計算書":  XMXA020_区分 = "2"
            Case "原価報告書":  XMXA020_区分 = "3"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               分類
    ' =========================================
    If p区分名 = "分類" Then
        Select Case pコード
            Case "分類1": XMXA020_区分 = "1"
            Case "分類2": XMXA020_区分 = "2"
            Case "分類3": XMXA020_区分 = "3"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               実績調整
    ' =========================================
    If p区分名 = "実績調整" Then
        Select Case pコード
            Case "全て有効":         XMXA020_区分 = "1"
            Case "本部経費振替のみ":  XMXA020_区分 = "2"
            Case "基幹データ調整のみ": XMXA020_区分 = "3"
            Case "全て無効":          XMXA020_区分 = "0"
            Case "基幹データ調整有効": XMXA020_区分 = "3" '基幹データ調整のみ
            Case "基幹データ調整無効": XMXA020_区分 = "0" '全て無効
            'Case "本部経費振替有効":  XMXA020_区分 = "2" '本部経費振替のみ
            'Case "本部経費振替無効":  XMXA020_区分 = "0" '全て無効
            Case Else: XMXA020_区分 = "0"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               販売売仕
    ' =========================================
    If p区分名 = "販売売仕" Then
        Select Case pコード
            Case "売上仕入有効": XMXA020_区分 = "1"
            Case "売上のみ有効": XMXA020_区分 = "2"
            Case "仕入のみ有効": XMXA020_区分 = "3"
            Case "売上仕入無効": XMXA020_区分 = "0"
            Case Else: XMXA020_区分 = "0"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               返済方法
    ' =========================================
    If p区分名 = "返済方法" Then
        Select Case pコード
            Case "元金均等返済": XMXA020_区分 = "1"
'            Case "元利均等返済": XMXA020_区分 = "2"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               借入貸付
    ' =========================================
    If p区分名 = "借入貸付" Then
        Select Case pコード
            Case "借入": XMXA020_区分 = "1"
            Case "貸付": XMXA020_区分 = "2"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               回収有無
    ' =========================================
    If p区分名 = "回収有無" Then
        Select Case pコード
            Case "回収無": XMXA020_区分 = "0"
            Case "回収有": XMXA020_区分 = "1"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               支払有無
    ' =========================================
    If p区分名 = "支払有無" Then
        Select Case pコード
            Case "支払無": XMXA020_区分 = "0"
            Case "支払有": XMXA020_区分 = "1"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               登録方法
    ' =========================================
    If p区分名 = "登録方法" Then
        Select Case pコード
            Case "標準登録": XMXA020_区分 = "0"
            Case "入力登録": XMXA020_区分 = "1"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               日割計算区分
    ' =========================================
    If p区分名 = "日割計算区分" Then
        Select Case pコード
            Case "自動計算": XMXA020_区分 = "0"
            Case "入力登録": XMXA020_区分 = "1"
        End Select
            
        Exit Function
    End If
'
    ' =========================================
    '               消費税率摘要区分
    ' =========================================
    If p区分名 = "消費税率摘要区分" Then
        Select Case pコード
            Case "契約時の消費税率": XMXA020_区分 = "0"
            Case "現行の消費税率": XMXA020_区分 = "1"
        End Select
            
        Exit Function
    End If
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
XMXA020_区分_ERR:
    pERR_MES = pPROGRAM_ID + "/ XMXA020_区分() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function



