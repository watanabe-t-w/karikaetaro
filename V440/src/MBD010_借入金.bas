Attribute VB_Name = "MBD010_借入金"
'
Option Explicit
'
Private Const pPROGRAM_ID As String = "MBD010_借入金"

'** MBD010_保証料算出 リターン **
Type MBD010_保証料算出リターン
    初回保証料 As Double
    保証料X年後(9) As Double
    解約保証料戻 As Double
End Type

Dim w保証料支払日 As Integer
Dim w内入開始年月日 As Variant      ' 08/12/05 V189
Dim w内入終了年月日 As Variant      ' 08/12/05 V189
Dim w利息対象内入開始年月日 As Variant      ' 08/12/05 V189
Dim w利息対象内入終了年月日 As Variant      ' 08/12/05 V189
Dim w融資残高 As Double             ' 08/12/05 V189
Dim w解約実行日 As Variant          '10/01/23
Dim w解約無効F As Integer           '10/01/24

Dim w借入内入 As MAA910_借入金内入
Dim wSM年月日(200) As Variant       '10/11/11 V189R
Dim wSM利率(200) As Double          '10/11/11 V189R
'
'------------------------------------------------
' MBD010_借入データセット
'------------------------------------------------
Public Function MBD010_借入データセット(pRs As ADODB.Recordset) As MAA910_借入金
'
    Dim j As Integer
    Dim w配列数 As Integer                                          ' 08/12/16 V189
    Dim X As Integer                                                ' 08/12/16 V189
    
    Dim Y As Integer                                                '10/11/11 V189R
    Dim w最終実績利率 As Double                                     '10/11/11 V189R
    Dim w最終実績年月日 As Variant                                  '10/11/11 V189R
    
    Dim w本日年月日 As Variant                                      '10/11/22 V189R
    
    Dim p金利SM利率 As MAA070_金利SM率
    Dim w借入金種別 As MAA070_借入金種別
'
    On Error GoTo MBD010_借入データセット_ERR
'
    w本日年月日 = Date                      '10/11/22 V189R
'
    MBD010_借入データセット.借入番号 = pRs("借入番号")
'
    '***借入金借換
    MBD010_借入データセット.保証会社区分 = P8.FCStr(pRs("保証会社区分"))
    MBD010_借入データセット.融資区分 = P8.FCStr(pRs("融資区分"))
'
    MBD010_借入データセット.プロジェクト番号 = P8.FCStr(pRs("プロジェクト番号"))
    MBD010_借入データセット.手入力区分 = pRs("手入力区分")          ' 07/02/10 V180
    MBD010_借入データセット.日割計算区分 = pRs("日割計算区分")
    
    MBD010_借入データセット.借入内容 = P8.FCStr(pRs("借入内容")) 'pRs("借入内容")
    MBD010_借入データセット.借入計画番号 = P8.FCStr(pRs("借入計画番号"))
    MBD010_借入データセット.金融リストラ番号 = P8.FCStr(pRs("金融リストラ番号"))
    MBD010_借入データセット.SM区分 = pRs("Sm区分")
    MBD010_借入データセット.銀行番号 = pRs("銀行番号")
    MBD010_借入データセット.支払日 = pRs("支払日")                  ' 07/01/30 V180
    MBD010_借入データセット.営業日区分 = pRs("営業日区分")          ' 07/01/30 V180
    MBD010_借入データセット.利息区分 = pRs("利息区分")              ' 07/01/30 V180
    MBD010_借入データセット.利息計算日数区分 = pRs("利息計算日数区分")  ' 07/01/30 V180
    MBD010_借入データセット.利息支払方法 = pRs("利息支払方法")      ' 07/01/30 V180
    MBD010_借入データセット.利息控除区分 = pRs("利息控除区分")      ' 07/01/30 V180
    MBD010_借入データセット.金利計算年間日数 = pRs("金利計算年間日数")  ' 07/01/30 V180
    MBD010_借入データセット.融資金額 = P8.FCDbl(pRs("融資金額"))
    MBD010_借入データセット.利率 = P8.FCDbl(pRs("利率"))
    MBD010_借入データセット.保証料率 = P8.FCDbl(pRs("保証料率"))
    MBD010_借入データセット.保証料分割フラグ = pRs("保証料分割フラグ")
    MBD010_借入データセット.実行日 = pRs("実行日")
    MBD010_借入データセット.初回返済年月 = pRs("初回返済年月")
    MBD010_借入データセット.初回返済実行日 = pRs("初回返済実行日")
    MBD010_借入データセット.金利初回年月 = pRs("金利初回年月")      'V180 07/02/01
    If IsNull(MBD010_借入データセット.金利初回年月) Then
        MBD010_借入データセット.金利初回年月 = pRs("初回返済年月")
    End If
    
    MBD010_借入データセット.最終返済年月 = pRs("最終返済年月")
    MBD010_借入データセット.最終返済実行日 = pRs("最終返済実行日")
    MBD010_借入データセット.解約年月 = pRs("解約年月")
    MBD010_借入データセット.解約実行日 = pRs("解約実行日")
            
    MBD010_借入データセット.解約保証料戻 = pRs("解約保証料戻")
    MBD010_借入データセット.金融解約年月 = pRs("金融解約年月")
    MBD010_借入データセット.金融解約実行日 = pRs("金融解約実行日")
    MBD010_借入データセット.金融解約保証料戻 = pRs("金融解約保証料戻")
                        
    MBD010_借入データセット.初回返済額 = P8.FCDbl(pRs("初回返済額"))
    MBD010_借入データセット.毎月返済額 = P8.FCDbl(pRs("毎月返済額"))
    MBD010_借入データセット.最終返済額 = P8.FCDbl(pRs("最終返済額"))
    MBD010_借入データセット.返済単位月数 = P8.FCDbl(pRs("返済単位月数"))
    MBD010_借入データセット.有担保フラグ = pRs("有担保フラグ")
    MBD010_借入データセット.担保名 = P8.FCStr(pRs("担保名"))
    MBD010_借入データセット.金利種別 = P8.FCDbl(pRs("金利種別"))    '08/01/20
    MBD010_借入データセット.金利条件 = P8.FCStr(pRs("金利条件"))    '08/01/20
    MBD010_借入データセット.基準金利区分 = P8.FCStr(pRs("基準金利区分"))
    MBD010_借入データセット.金利グループ区分 = P8.FCStr(pRs("金利グループ区分"))
    
    MBD010_借入データセット.長短区分 = P8.FCDbl(pRs("長短区分"))
    MBD010_借入データセット.設備フラグ = pRs("設備フラグ")
    MBD010_借入データセット.資金用途 = P8.FCStr(pRs("資金用途"))
    MBD010_借入データセット.自己資金フラグ = pRs("自己資金フラグ")
    MBD010_借入データセット.据置回数 = pRs("据置回数")
    MBD010_借入データセット.支払回数 = pRs("支払回数")
    
    MBD010_借入データセット.借入貸付 = pRs("借入貸付")      '06/03/27 V150
    MBD010_借入データセット.借入金種別区分 = P8.FCStr(pRs("借入金種別区分"))
    MBD010_借入データセット.返済方法 = pRs("返済方法")      '06/03/27 V150
    
    w借入金種別 = MAA070_借入金種別Read(MBD010_借入データセット.借入金種別区分)
    MBD010_借入データセット.社債フラグ = w借入金種別.社債フラグ
    MBD010_借入データセット.利子補給金フラグ = w借入金種別.利子補給金フラグ '16/03/26 利子補給に伴う変更
    
    '***TEST 用 残高推移表CHECKのため
    'If MBD010_借入データセット.金利グループ区分 = "" Then
    '    MBD010_借入データセット.金利グループ区分 = "10"
    'End If
                    
    '仮セット
    MBD010_借入データセット.変動最終利率 = -1
        
    For j = 2 To 100                                        ' 07/07/04 V188
        MBD010_借入データセット.金利(j).金利変更x回目年月 = pRs("金利変更" + CStr(j) + "回目年月")
        MBD010_借入データセット.金利(j).金利x回目 = P8.FCDbl(pRs("金利" + CStr(j) + "回目"))
        
        If P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) = MBD010_借入データセット.金利種別 _
        And Not IsNull(MBD010_借入データセット.金利(j).金利変更x回目年月) Then
            MBD010_借入データセット.変動最終利率 = MBD010_借入データセット.金利(j).金利x回目
        End If
    Next
    
    MBD010_借入データセット.融資可能枠 = pRs("融資可能枠")
    MBD010_借入データセット.融資残高 = pRs("融資残高")
    MBD010_借入データセット.借入年度 = pRs("借入年度")
    
    MBD010_借入データセット.取消フラグ = pRs("取消フラグ")
'
    '***借入金内入
    Call MBD010_借入内入_クリア
    
    w借入内入 = MBD010_借入内入(MBD010_借入データセット.借入番号)
'
    '----------< 金利シミュレーション >----------------------------------------
    'If G金利SM = True Then
    '    p金利SM利率 = MAA070_金利SM率Read(MBD010_借入データセット.金利グループ区分)
    '    For X = 1 To 100
    '        GStr = p金利SM利率.金利グループ区分
    '        GStr = p金利SM利率.利率増減率(X).年月日
    '        GStr = p金利SM利率.利率増減率(X).増減率
    '    Next X
    'End If
     
'
    '***金利変更X回目年月の調整（年月　TO　年月日)
    w解約無効F = 1                                                  '10/01/24
    Call MBD010_借入金テーブル作成("", MBD010_借入データセット)     ' 08/12/16 V189
    w解約無効F = 0                                                  '10/01/24

    w配列数 = UBound(G借入金テーブル)                               ' 08/12/16 V189
    For j = 2 To 100                                                ' 08/12/16 V189
      If IsNull(MBD010_借入データセット.金利(j).金利変更x回目年月) Then '10/01/24
        Exit For                                                    '10/11/11 N189R
      End If                                                        '10/01/24
        For X = 1 To w配列数                                        ' 08/12/16 V189
            If (G借入金テーブル(X).据置x回目 = 2 Or G借入金テーブル(X).据置x回目 = 4) _
                And G借入金テーブル(X).日割日数 > 0 _
                And Format(G借入金テーブル(X).返済予定年月, "yyyy/mm/dd") _
                    = Format(MBD010_借入データセット.金利(j).金利変更x回目年月, "yyyy/mm/dd") Then
                     
                '初回返済年月日ＯＲ　最終返済年月日　が　手打ちの場合
                G借入金テーブル(X).利息計算年月日 = MBD010_利息計算年月日(G借入金テーブル(X).返済予定年月, _
                                               MBD010_借入データセット.支払日, _
                                               MBD010_借入データセット.営業日区分, _
                                               MBD010_借入データセット.利息計算日数区分)    '10/01/03
                                               
                If Format(G借入金テーブル(X).返済予定年月, "yyyy/mm/dd") = _
                   Format(MBD010_借入データセット.初回返済年月, "yyyy/mm/dd") _
                   And Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                       Format(MBD010_借入データセット.初回返済実行日, "yyyy/mm/dd") Then    '10/01/03
                   G借入金テーブル(X).実際年月日 = MBD010_借入データセット.初回返済実行日   '10/01/03
                   G借入金テーブル(X).利息計算年月日 = MBD010_借入データセット.初回返済実行日   '10/01/23
                Else                                                                    '10/01/03
                    If Format(G借入金テーブル(X).返済予定年月, "yyyy/mm/dd") = _
                       Format(MBD010_借入データセット.最終返済年月, "yyyy/mm/dd") _
                       And Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                           Format(MBD010_借入データセット.最終返済実行日, "yyyy/mm/dd") Then    '10/01/03       Then  '10/01/03
                       G借入金テーブル(X).実際年月日 = MBD010_借入データセット.最終返済実行日   '10/01/03
                       G借入金テーブル(X).利息計算年月日 = MBD010_借入データセット.最終返済実行日
                    Else                                                                '10/01/03
                   '利息計算固定日数 10/01/03 削除
                   '   If MBD010_借入データセット.利息計算日数区分 = 1 Then
                   '     G借入金テーブル(X).実際年月日 = C年月日.GetDate("設定", _
                   '             G借入金テーブル(X).返済予定年月, MBD010_借入データセット.支払日)
                   '   End If
                    End If                                                              '10/01/03
                End If                                                                  '10/01/03
                
                '**解約日の時調整
                'If Format(MBD010_借入データセット.解約実行日, "yyyy/mm/dd") = _
                '        Format(G借入金テーブル(X).実際年月日, "yyyy/mm/dd") Then            '10/01/23
                '    G借入金テーブル(X).利息計算年月日 = MBD010_借入データセット.解約実行日  '10/01/23
                'End If                                                                      '10/01/23
                
                
                If MBD010_借入データセット.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
                    '利息先払
                    MBD010_借入データセット.金利(j).金利変更x回目年月 = G借入金テーブル(X).利息計算年月日   '10/01/23
                Else
                    '利息後払
                    MBD010_借入データセット.金利(j).金利変更x回目年月 = _
                        DateAdd("d", -G借入金テーブル(X).利息対象期間日数, G借入金テーブル(X).利息計算年月日)
                End If
                GoTo next2
            End If
        Next
        
next2:
    
    Next
    
    
    '***金利シュミュレーションテーブルセット1
    
    For j = 0 To 200                    '10/11/11 V189R
        wSM年月日(j) = Null             '10/11/11 V189R
        wSM利率(j) = 0                  '10/11/11 V189R
    Next                                '10/11/11 V189R
    '***仮にSET
'    G金利SM = True
    
    '*標準入力＆変動金利＆未来利率シュミュレーション
    If G金利SM = True Then              '10/11/11 V189R
        p金利SM利率 = MAA070_金利SM率Read(MBD010_借入データセット.金利グループ区分)  '10/11/11 V189R
    End If                                                                           '10/11/11 V189R
    
    If MBD010_借入データセット.手入力区分 = 0 And MBD010_借入データセット.金利種別 = 0 _
       And G金利SM = True Then          '10/11/11 V189R
       
       For j = 2 To 100                 '10/11/11 V189R
            If IsNull(MBD010_借入データセット.金利(j).金利変更x回目年月) Then   '10/11/11 V189R
                Exit For                                                        '10/11/11 V189R
            Else                                                                '10/11/11 V189R
                wSM年月日(j) = MBD010_借入データセット.金利(j).金利変更x回目年月    '10/11/11 V189R
                wSM利率(j) = MBD010_借入データセット.金利(j).金利x回目          '10/11/11 V189R
            End If                                                              '10/11/11 V189R
       Next                                                                     '10/11/11 V189R
       
       If j = 2 Then                                                            '10/11/11 V189R
            w最終実績年月日 = MBD010_借入データセット.実行日                    '10/11/11 V189R
            w最終実績利率 = MBD010_借入データセット.利率                        '10/11/11 V189R
       Else                                                                     '10/11/11 V189R
            If j = 100 Then                                                     '10/11/11 V189R
                w最終実績年月日 = MBD010_借入データセット.金利(j).金利変更x回目年月 '10/11/11 V189R
                w最終実績利率 = MBD010_借入データセット.金利(j).金利x回目       '10/11/11 V189R
            Else                                                                '10/11/11 V189R
                w最終実績年月日 = MBD010_借入データセット.金利(j - 1).金利変更x回目年月 '10/11/11 V189R
                w最終実績利率 = MBD010_借入データセット.金利(j - 1).金利x回目   '10/11/11 V189R
            End If                                                              '10/11/11 V189R
       End If                                                                   '10/11/11 V189R
       
       '**金利シュミュレーションテーブルセット2
       
       For X = 1 To 100                                                         '10/11/11 V189R
            If IsNull(p金利SM利率.利率増減率(X).年月日) Then                    '10/11/11 V189R
                Exit For                                                        '10/11/11 V189R
            Else                                                                '10/11/11 V189R
                If Format(p金利SM利率.利率増減率(X).年月日, "yyyy/mm/dd") <= Format(w最終実績年月日, "yyyy/mm/dd") _
                   Or Format(MBD010_借入データセット.実行日, "yyyy/mm/dd") >= Format(p金利SM利率.利率増減率(X).年月日, "yyyy/mm/dd") _
                   Or Format(w本日年月日, "yyyy/mm/dd") >= Format(p金利SM利率.利率増減率(X).年月日, "yyyy/mm/dd") Then    '10/11/22 V189R
                    GoTo Ok1                                                    '10/11/11 V189R
                Else                                                            '10/11/11 V189R
                
                    '*直近の支払年月日を算出
                    w配列数 = UBound(G借入金テーブル)                               '10/11/11 V189R
                    For Y = 1 To w配列数                                            '10/11/11 V189R
                        If Format(p金利SM利率.利率増減率(X).年月日, "yyyy/mm/dd") <= Format(w最終実績年月日, "yyyy/mm/dd") Then '10/11/11 V189R
                            GoTo Ok1                                                '10/11/11 V189R
                        Else                                                        '10/11/11 V189R
                            If Format(G借入金テーブル(Y).利息計算年月日, "yyyy/mm/dd") < _
                                Format(p金利SM利率.利率増減率(X).年月日, "yyyy/mm/dd") Then

                                GoTo Ok2                                            '10/11/11 V189R
                            Else                                                    '10/11/11 V189R
                            
                                'If G借入金テーブル(Y).利息額 = 0 Then               '10/11/11 V189R
                                '    GoTo Ok2                                        '10/11/13 V189R
                                'End If                                              '10/11/13 V189R
                                
                                '*同一の次回支払年月日に対して、金利変更が複数回存在した場合の対応
                                If Format(G借入金テーブル(Y).利息計算年月日, "yyyy/mm/dd") = _
                                    Format(wSM年月日(j - 1), "yyyy/mm/dd") Then     '10/11/13 V189R
                                    
                                    wSM利率(j - 1) = p金利SM利率.利率増減率(X).増減率 + w最終実績利率 '10/11/13 V189R
                                Else                                                '10/11/13 V189R
                                    
                                    wSM年月日(j) = G借入金テーブル(Y).利息計算年月日    '10/11/11 V189R
                                    wSM利率(j) = p金利SM利率.利率増減率(X).増減率 + w最終実績利率   '10/11/11 V189R
                            
                                    j = j + 1                                       '10/11/11 V189R
                                End If                                              '10/11/13 V189R
                                
                            End If                                                  '10/11/11 V189R
                        
                            Exit For                                                '10/11/11 V189R
                        End If                                                      '10/11/11 V189R
Ok2:                                                                                '10/11/11 V189R

                    Next                                                            '10/11/11 V189R
                    
                End If                                                          '10/11/11 V189R
                
            End If                                                                      '10/11/11 V189R
Ok1:
            
       Next                                                                     '10/11/11 V189R
       
    End If                                                                      '10/11/11 V189R
    
       
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_借入データセット_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入データセット() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

'------------------------------------------------
' MBD010_借入内入_クリア
'------------------------------------------------
Public Sub MBD010_借入内入_クリア()
'
    Dim j As Integer
'
    w借入内入.内入区分 = False

    For j = 0 To 500
        w借入内入.内入(j).内入x回目年月日 = Null
        w借入内入.内入(j).内入金額x回目 = 0
        w借入内入.内入(j).手数料x回目 = 0
    Next j
'
End Sub

'------------------------------------------------
' MBD010_借入内入
'------------------------------------------------
Public Function MBD010_借入内入(pNo As String) As MAA910_借入金内入
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String

    Dim j As Integer
'
    On Error GoTo MBD010_借入内入_ERR
'
    w借入内入.内入区分 = False

    wstr = ""
    wstr = wstr + "Select * From DBDA010_借入金内入1"
    wstr = wstr + " Where 借入番号='" & pNo & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
        For j = 1 To 80
            If IsNull(wRs("内入" + CStr(j) + "回目年月日")) Then
                Exit For
            End If
            
            MBD010_借入内入.内入(j).内入x回目年月日 = wRs("内入" + CStr(j) + "回目年月日")
            MBD010_借入内入.内入(j).内入金額x回目 = P8.FCDbl(wRs("内入金額" + CStr(j) + "回目"))
            MBD010_借入内入.内入(j).手数料x回目 = P8.FCDbl(wRs("手数料" + CStr(j) + "回目"))
        
            w借入内入.内入区分 = True
        Next
    End If
    
    wRs.Close
    Set wRs = Nothing
'
    wstr = ""
    wstr = wstr + "Select * From DBDA010_借入金内入2"
    wstr = wstr + " Where 借入番号='" & pNo & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
        For j = 81 To 160
            If IsNull(wRs("内入" + CStr(j) + "回目年月日")) Then
                Exit For
            End If
            
            MBD010_借入内入.内入(j).内入x回目年月日 = wRs("内入" + CStr(j) + "回目年月日")
            MBD010_借入内入.内入(j).内入金額x回目 = P8.FCDbl(wRs("内入金額" + CStr(j) + "回目"))
            MBD010_借入内入.内入(j).手数料x回目 = P8.FCDbl(wRs("手数料" + CStr(j) + "回目"))
        
            w借入内入.内入区分 = True
        Next
    End If
    
    wRs.Close
    Set wRs = Nothing
'
    wstr = ""
    wstr = wstr + "Select * From DBDA010_借入金内入3"
    wstr = wstr + " Where 借入番号='" & pNo & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
        For j = 161 To 240
            If IsNull(wRs("内入" + CStr(j) + "回目年月日")) Then
                Exit For
            End If
            
            MBD010_借入内入.内入(j).内入x回目年月日 = wRs("内入" + CStr(j) + "回目年月日")
            MBD010_借入内入.内入(j).内入金額x回目 = P8.FCDbl(wRs("内入金額" + CStr(j) + "回目"))
            MBD010_借入内入.内入(j).手数料x回目 = P8.FCDbl(wRs("手数料" + CStr(j) + "回目"))
        
            w借入内入.内入区分 = True
        Next
    End If
    
    wRs.Close
    Set wRs = Nothing
'
    wstr = ""
    wstr = wstr + "Select * From DBDA010_借入金内入4"
    wstr = wstr + " Where 借入番号='" & pNo & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
        For j = 241 To 320
            If IsNull(wRs("内入" + CStr(j) + "回目年月日")) Then
                Exit For
            End If
            
            MBD010_借入内入.内入(j).内入x回目年月日 = wRs("内入" + CStr(j) + "回目年月日")
            MBD010_借入内入.内入(j).内入金額x回目 = P8.FCDbl(wRs("内入金額" + CStr(j) + "回目"))
            MBD010_借入内入.内入(j).手数料x回目 = P8.FCDbl(wRs("手数料" + CStr(j) + "回目"))
        
            w借入内入.内入区分 = True
        Next
    End If
    
    wRs.Close
    Set wRs = Nothing
'
    wstr = ""
    wstr = wstr + "Select * From DBDA010_借入金内入5"
    wstr = wstr + " Where 借入番号='" & pNo & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
        For j = 321 To 400
            If IsNull(wRs("内入" + CStr(j) + "回目年月日")) Then
                Exit For
            End If
            
            MBD010_借入内入.内入(j).内入x回目年月日 = wRs("内入" + CStr(j) + "回目年月日")
            MBD010_借入内入.内入(j).内入金額x回目 = P8.FCDbl(wRs("内入金額" + CStr(j) + "回目"))
            MBD010_借入内入.内入(j).手数料x回目 = P8.FCDbl(wRs("手数料" + CStr(j) + "回目"))
        
            w借入内入.内入区分 = True
        Next
    End If
    
    wRs.Close
    Set wRs = Nothing
'
    wstr = ""
    wstr = wstr + "Select * From DBDA010_借入金内入6"
    wstr = wstr + " Where 借入番号='" & pNo & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
        For j = 401 To 480
            If IsNull(wRs("内入" + CStr(j) + "回目年月日")) Then
                Exit For
            End If
            
            MBD010_借入内入.内入(j).内入x回目年月日 = wRs("内入" + CStr(j) + "回目年月日")
            MBD010_借入内入.内入(j).内入金額x回目 = P8.FCDbl(wRs("内入金額" + CStr(j) + "回目"))
            MBD010_借入内入.内入(j).手数料x回目 = P8.FCDbl(wRs("手数料" + CStr(j) + "回目"))
        
            w借入内入.内入区分 = True
        Next
    End If
    
    wRs.Close
    Set wRs = Nothing
'
    wstr = ""
    wstr = wstr + "Select * From DBDA010_借入金内入7"
    wstr = wstr + " Where 借入番号='" & pNo & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
        For j = 481 To 500
            If IsNull(wRs("内入" + CStr(j) + "回目年月日")) Then
                Exit For
            End If
            
            MBD010_借入内入.内入(j).内入x回目年月日 = wRs("内入" + CStr(j) + "回目年月日")
            MBD010_借入内入.内入(j).内入金額x回目 = P8.FCDbl(wRs("内入金額" + CStr(j) + "回目"))
            MBD010_借入内入.内入(j).手数料x回目 = P8.FCDbl(wRs("手数料" + CStr(j) + "回目"))
        
            w借入内入.内入区分 = True
        Next
    End If
    
    wRs.Close
    Set wRs = Nothing
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_借入内入_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入内入() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)
    
    End
'
End Function

'------------------------------------------------
' MBD010_借入明細作成
'------------------------------------------------
Public Sub MBD010_借入明細作成(p金融リストラ番号 As String, p借入計画マスタ As MAA910_借入金)
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String

    Dim j As Integer
    Dim w解約実行日 As Variant      ' 07/02/21 V180
    Dim w返済回数 As Integer        '10/02/27
'
    On Error GoTo MBD010_借入明細作成_ERR
'
    w返済回数 = 0                   '10/02/27
'
    wstr = ""
    wstr = wstr + "Select * From DCDA020_借入金明細"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        For j = 1 To UBound(G借入金テーブル)
          If G借入金テーブル(j).元金額 <> 0 Or G借入金テーブル(j).利息額 <> 0 _
             Or (G借入金テーブル(j).融資残高 <> 0 And p借入計画マスタ.実行日 = G借入金テーブル(j).実際年月日) _
             Or G借入金テーブル(j).保証料 <> 0 Or G借入金テーブル(j).手数料 <> 0 _
             Or Format(w解約実行日, "yyyymmdd") = _
                   Format(G借入金テーブル(j).実際年月日, "yyyymmdd") Then '10/06/16 V195
             
            wRs.AddNew
        
                wRs("借入番号") = G借入金テーブル(j).借入番号
                If (Format(p借入計画マスタ.実行日, "yyyy/mm/dd") _
                        <> Format(G借入金テーブル(j).実際年月日, "yyyy/mm/dd")) _
                   Or (p借入計画マスタ.実行日 = G借入金テーブル(j).実際年月日 _
                       And p借入計画マスタ.実行日 = p借入計画マスタ.初回返済実行日) Then  '10/05/06 V195
                    w返済回数 = w返済回数 + 1                   '10/02/27
                    wRs("返済回数") = w返済回数                 '10/02/27
                Else                                            '10/02/27
                    wRs("返済回数") = 0                         '10/02/27
                End If                                          '10/02/27
                'wRs("返済回数") = G借入金テーブル(j).返済回数
                wRs("据置X回目") = G借入金テーブル(j).据置x回目
                wRs("返済予定年月") = G借入金テーブル(j).返済予定年月
    
                wRs("実際年月日") = G借入金テーブル(j).実際年月日
                wRs("利息計算年月日") = G借入金テーブル(j).利息計算年月日   '10/01/04
                
                'If p借入計画マスタ.借入金種別区分 = "04" Then                    '16/03/23
                    
                '    wRs("利息額") = -G借入金テーブル(j).利息額    '16/03/23
                    
                '    wRs("返済金額") = -G借入金テーブル(j).利息額 + G借入金テーブル(j).元金額 '16/03/23
                    
                'Else                                             '16/03/23
                '    wRs("利息額") = G借入金テーブル(j).利息額    '16/03/23
                'End If                                            '16/03/23
                
                
                wRs("保証料") = G借入金テーブル(j).保証料
                wRs("手数料") = G借入金テーブル(j).手数料       ' 08/12/06 V189
                wRs("金融保証料") = G借入金テーブル(j).金融保証料
                
                '解約算出
                If p金融リストラ番号 > "" _
                   And p金融リストラ番号 = p借入計画マスタ.金融リストラ番号 Then  ' 07/02/18 V180
                    w解約実行日 = p借入計画マスタ.金融解約実行日
                Else
                    w解約実行日 = p借入計画マスタ.解約実行日
                End If
                
                If Format(w解約実行日, "yyyymmdd") = _
                   Format(G借入金テーブル(j).実際年月日, "yyyymmdd") Then   ' 07/02/21 V180
                    wRs("返済金額") = G借入金テーブル(j).融資残高 + G借入金テーブル(j).利息額
                    wRs("元金額") = G借入金テーブル(j).融資残高             ' 07/02/21 V180
                    wRs("融資残高") = 0                                     ' 07/02/21 V180
                Else                                                        ' 07/02/21 V180
                    wRs("返済金額") = G借入金テーブル(j).返済金額
                    wRs("元金額") = G借入金テーブル(j).元金額
                    wRs("融資残高") = G借入金テーブル(j).融資残高
                End If                                                      ' 07/02/21 V180
                
                'If p借入計画マスタ.借入金種別区分 = "04" Then                    '16/03/23
                If p借入計画マスタ.利子補給金フラグ = 1 Then                    '16/03/23 利子補給に伴う変更
                   
                    wRs("利息額") = -G借入金テーブル(j).利息額    '16/03/23
                    
                    wRs("返済金額") = -G借入金テーブル(j).利息額 + G借入金テーブル(j).元金額 '16/03/23
                    
                Else                                             '16/03/23
                    wRs("利息額") = G借入金テーブル(j).利息額    '16/03/23
                End If                                           '16/03/23
                
                 
                    
                wRs("日割日数") = G借入金テーブル(j).日割日数
                wRs("利息対象期間日数") = G借入金テーブル(j).利息対象期間日数       'V182 2008/01/28
                wRs("利率") = G借入金テーブル(j).利率
                
            wRs.Update
          End If                    '10/02/27
          
        Next
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_借入明細作成_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入明細作成() でエラー" + vbCrLf + vbCrLf + _
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
' MBD010_借入明細作成_入力登録
'------------------------------------------------
Public Sub MBD010_借入明細作成_入力登録(p借入計画マスタ As MAA910_借入金)
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String

    Dim j As Integer
'
    On Error GoTo MBD010_借入明細作成_入力登録_ERR
'
    wstr = ""
    wstr = wstr + "Select * From DCDA020_借入金明細"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        For j = 1 To UBound(G借入金入力)
             
            wRs.AddNew
        
                wRs("借入番号") = p借入計画マスタ.借入番号
                wRs("返済回数") = j
                wRs("返済予定年月") = G借入金入力(j).借入返済年月日
                wRs("実際年月日") = G借入金入力(j).借入返済年月日
                wRs("利息計算年月日") = G借入金入力(j).利息計算年月日
                
                wRs("元金額") = G借入金入力(j).元金
                
                'If p借入計画マスタ.借入金種別区分 = "04" Then           '16/03/23
                If p借入計画マスタ.利子補給金フラグ = 1 Then           '16/03/23 利子補給に伴う変更
                    wRs("利息額") = -G借入金入力(j).利息額              '16/03/23
                    wRs("返済金額") = -G借入金入力(j).利息額 + G借入金入力(j).元金  '16/03/23
                Else                                                    '16/03/23
                    wRs("利息額") = G借入金入力(j).利息額               '16/03/23
                    wRs("返済金額") = G借入金入力(j).返済金額           '16/03/23
                End If                                                  '16/03/23
                
                wRs("融資残高") = G借入金入力(j).融資残高
                wRs("仮計上利息額") = G借入金入力(j).仮計上利息額
                
                wRs("日割日数") = G借入金入力(j).日割日数
                wRs("利息対象期間日数") = G借入金入力(j).利息対象期間日数
                wRs("利率") = G借入金入力(j).利率
                
                wRs("据置X回目") = Null
                'wRs("保証料") = Null
                wRs("金融保証料") = Null
            
                wRs("初期手数料") = G借入金入力(j).初期手数料
                wRs("元金手数料") = G借入金入力(j).元金手数料
                wRs("利息手数料") = G借入金入力(j).利息手数料
                wRs("保証料") = G借入金入力(j).保証料
            
            wRs.Update
          
        Next
    
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_借入明細作成_入力登録_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入明細作成_入力登録() でエラー" + vbCrLf + vbCrLf + _
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
' MBD010_借入明細作成_明細TR
'------------------------------------------------
Public Sub MBD010_借入明細作成_明細TR(pKar As String, pTbl As String)
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String, ws01 As String, ws02 As String
'
    On Error GoTo MBD010_借入明細作成_明細TR_ERR
'
    ws01 = ""
    ws02 = ""
    
    ws01 = ws01 & "借入番号,"
    ws01 = ws01 & "返済回数,"
    ws01 = ws01 & "返済予定年月,"
    ws01 = ws01 & "実際年月日,"
    ws01 = ws01 & "利息計算年月日,"
    ws01 = ws01 & "返済金額,"
    ws01 = ws01 & "元金額,"
    ws01 = ws01 & "融資残高,"
    ws01 = ws01 & "利息額,"
    ws01 = ws01 & "仮計上利息額,"
    ws01 = ws01 & "日割日数,"
    ws01 = ws01 & "利息対象期間日数,"
    ws01 = ws01 & "利率,"
    
    ws02 = ws02 & "据置X回目,"
    ws02 = ws02 & "保証料,"
    ws02 = ws02 & "金融保証料"
    
    wstr = ""
    wstr = wstr & "Insert Into DCDA020_借入金明細"
    wstr = wstr & "(" & ws01 & ws02 & ")"
    wstr = wstr & " Select "
    
    wstr = wstr & "KM.借入番号,"
    wstr = wstr & "KM.返済回数,"
    wstr = wstr & "KM.返済予定年月,"
    wstr = wstr & "KM.実際年月日,"
    wstr = wstr & "KM.利息計算年月日,"
    wstr = wstr & "KM.返済金額,"
    wstr = wstr & "KM.元金額,"
    wstr = wstr & "KM.融資残高,"
    wstr = wstr & "KM.利息額,"
    wstr = wstr & "KM.仮計上利息額,"
    wstr = wstr & "KM.日割日数,"
    wstr = wstr & "KM.利息対象期間日数,"
    wstr = wstr & "KM.利率,"

    wstr = wstr & "null,null,null"
    wstr = wstr & " From " & pTbl & " As KM"
    wstr = wstr & " INNER JOIN DBDA010_借入金 As K"
    wstr = wstr & " ON KM.借入番号 = K.借入番号"
    
    If pKar <> "" Then
        wstr = wstr & " Where KM.借入番号='" & pKar & "'"
        wstr = wstr & " AND KM.取消フラグ=0"
        wstr = wstr & " AND KM.取消フラグ２=0"
        wstr = wstr & " AND K.手入力区分=1"
    Else
        wstr = wstr & " WHERE KM.取消フラグ=0"
        wstr = wstr & " AND KM.取消フラグ２=0"
        wstr = wstr & " AND K.手入力区分=1"
    End If
    
    GDb.Execute wstr
    '
    DoEvents

    wstr = ""
    wstr = wstr + "Select last(利率) As 最新利率 From DCDA020_借入金明細"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
        GDbl2 = P8.FCDbl(wRs("最新利率"))
    End If
    wRs.Close
    Set wRs = Nothing
    
    wstr = ""
    wstr = wstr + "Select count(借入番号) As 支払回数 From DCDA020_借入金明細"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
        GInt1 = P8.FCDbl(wRs("支払回数"))
    End If
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_借入明細作成_明細TR_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入明細作成_明細TR() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

''------------------------------------------------
'' MBD010_借入明細作成_明細TR_指定番号年月
''------------------------------------------------
'Public Function MBD010_借入明細作成_明細TR_指定番号年月(pNengetu As Date)
''
'    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
'    Dim wstr As String, wstr2 As String
'
'    Dim ws01 As String
'    Dim wdate As Date
''
'    On Error GoTo MBD010_借入明細作成_明細TR_指定番号年月_ERR
''
'    ws01 = ""
'    wdate = DateAdd("m", 1, pNengetu)
'
'    wstr2 = ""
'    wstr2 = wstr2 & "Select * From DBDA010_借入金明細TR"
'    wstr2 = wstr2 & " Order by 実際年月日 desc"
'    Call AdoRecordsetOpen(GDb, wRs2, wstr2)
'
'    wstr = ""
'    wstr = wstr + "Select * From DCDA020_借入金明細"
'    Call AdoRecordsetOpen(GDb, wRs, wstr)
'
'    Do Until wRs2.EOF
'        If ws01 <> wRs2("借入番号") _
'        And Format(wRs2("実際年月日"), "yyyy/mm/dd") < Format(wdate, "yyyy/mm/dd") Then
'
'            wRs.AddNew
'
'            wRs("借入番号") = wRs2("借入番号")
'            wRs("返済回数") = wRs2("返済回数")
'            wRs("据置X回目") = wRs2("据置X回目")
'            wRs("返済予定年月") = wRs2("返済予定年月")
'            wRs("実際年月日") = wRs2("実際年月日")
'            wRs("返済金額") = wRs2("返済金額")
'            wRs("元金額") = wRs2("元金額")
'            wRs("利息額") = wRs2("利息額")
'            wRs("保証料") = wRs2("保証料")
'            wRs("金融保証料") = wRs2("金融保証料")
'            wRs("手数料") = wRs2("手数料")
'            wRs("融資残高") = wRs2("融資残高")
'            wRs("日割日数") = wRs2("日割日数")
'            wRs("利率") = wRs2("利率")
'
'            wRs.Update
'
'            ws01 = wRs2("借入番号")
'
'        End If
'
'        wRs2.MoveNext
'    Loop
'
'    wRs.Close
'    Set wRs = Nothing
''
'    Exit Function
''
''----------< ERROR ROUTINE >---------------------------------------------------
'MBD010_借入明細作成_明細TR_指定番号年月_ERR:
'    pERR_MES = pPROGRAM_ID + "/ MBD010_借入明細作成_明細TR_指定番号年月() でエラー" + vbCrLf + vbCrLf + _
'                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
'                "プロジェクト名：" + Err.Source + vbCrLf + _
'                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
'                GProduct + "を終了します"
'    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
'    pERR_RET = PUT_LOG(pERR_MES)
'
'    End
''
'End Function

'------------------------------------------------
' MBD010_借入金テーブル作成
'------------------------------------------------
Public Function MBD010_借入金テーブル作成(p金融リストラ As String, p借入金マスタ As MAA910_借入金, _
                                        Optional pClear As Boolean = True) As Double
'
    Dim j As Integer

    Dim wTable As MAA910_借入金テーブル
    
    Dim w解約保証料戻

    Dim w銀行マスタ As MAA030_銀行

    Dim w実行支払年月 As Date
    Dim w支払日 As Integer
    Dim w返済予定年月 As Date
    Dim w実際年月日 As Date
    Dim wRecCount As Integer
    
    Dim w実行解約年月 As Date
    
    'Dim w融資残高 As Double
    
    Dim wCount As Integer
    Dim wSm解約実行日 As Date
    Dim w実際予定年月日 As Date
    Dim wRet保証料算出 As MBD010_保証料算出リターン
    
    'Dim w解約実行日 As Variant
    Dim w解約年月 As Variant                    ' 08/07/20 V188
    Dim w年月日 As Date
    
    Dim w返済単位判定 As Long
    Dim w返済単位回数 As Integer
    
    Dim w据置回数 As Integer
    
    Dim wp日数 As Integer
    Dim wd01 As Date
    Dim w調整月数 As Integer    'V180 07/02/01
    Dim w日割日数   As Integer        ' 07/03/23 V180
    Dim w利息対象期間日数   As Integer      'V182 2008/01/28
    Dim OLD実際年月日 As Variant    ' 07/03/24 V180
    Dim NEW実際年月日 As Variant    ' 07/03/24 V180
    Dim w差月数 As Integer          ' 07/03/24 V180
    
    Dim w一括支払利息開始年月 As Date           '08/03/14 V185
    Dim wGAP As Integer                         '08/03/14 V185
    Dim w初回利息F As Integer                   '08/03/14 V185
    Dim w有効F As Integer                       '08/03/14 V185
    
    Dim sv支払回数 As Integer                   ' 08/07/19 V188
    Dim w調整解約年月 As Date                   ' 08/07/20 V188
    Dim wyy As Integer                          ' 08/07/20 V188
    Dim wmm As Integer                          ' 08/07/20 V188
    Dim wdd As Integer                          ' 08/07/20 V188
    
    Dim w本日年月日 As Date                     ' 08/12/06 V189
    Dim w利息計算基準年月日 As Date             ' 08/12/10 V189
    Dim w実行日 As Date                         '10/01/21
    
    Dim wx As Integer                           '10/02/05
    Dim wY As Integer                           '10/02/05
    Dim wZ As Integer                           '10/02/05
    Dim wW As Integer                           '10/02/05
'
    On Error GoTo MBD010_借入金テーブル作成_ERR
'
    w本日年月日 = Date                          ' 08/12/06 V189
    
    wp日数 = 0
    
        
    '*********
    '*
    '*           w解約実行日　セット
    '*
    '*********
    w解約実行日 = Null
    If w解約無効F = 0 Then                      '10/01/24
        If p金融リストラ <> "" And p金融リストラ = p借入金マスタ.金融リストラ番号 Then
            w解約実行日 = p借入金マスタ.金融解約実行日
            w解約年月 = p借入金マスタ.金融解約年月              ' 08/07/20 V188
        Else
            If Not IsNull(p借入金マスタ.解約実行日) Then
                w解約実行日 = p借入金マスタ.解約実行日
                w解約年月 = p借入金マスタ.解約年月              ' 08/07/20 V188
            End If
        End If
    End If                                                      '10/01/24
    
    w銀行マスタ = MAA030_銀行マスタRead(p借入金マスタ.銀行番号)
    
    
    
    '実行支払年月 の　算出      10/01/21
    w実行支払年月 = MBD010_実行日支払年月算出(p借入金マスタ.実行日, p借入金マスタ.営業日区分, _
                                           p借入金マスタ.支払日)
    
    'w支払日 = Day(C年月日.GetDate("月末", p借入金マスタ.実行日))
    
    'w実行支払年月 = Format(p借入金マスタ.実行日, "yyyy/mm/01")  '10/01/16
    'If p借入金マスタ.営業日区分 = 0 Then                           '10/01/20
    '    wyy = Year(p借入金マスタ.実行日)                            '10/01/20
    '    wmm = Month(p借入金マスタ.実行日)                           '10/01/20
    '    w実行支払年月 = Right("0000" & CStr(wyy), 4) & "/" & Right("00" & CStr(wmm), 2) & "/" & Right("00" & CStr(w支払日), 2)
    '    w実行支払年月 = Format(w実行支払年月, "yyyy/mm/dd")         '10/01/20
    'Else                                                            '10/01/20
    '    w実行支払年月 = MXA030_翌営業年月日計算(w実行支払年月, p借入金マスタ.支払日, p借入金マスタ.営業日区分)  '10/01/20
    'End If                                                          '10/01/20
    
    
    'If Format(w実行支払年月, "yyyy/mm/dd") <= Format(p借入金マスタ.実行日, "yyyy/mm/dd") Then '10/01/16
    '    w実行支払年月 = DateAdd("m", 1, w実行支払年月)                                          '10/01/16
    'End If                  '10/01/16
    'w実行支払年月 = Format(w実行支払年月, "yyyy/mm/01")     '10/01/16
    
     
    
    
    '**********
    '*
    '*           据置回数&支払回数　再セット
    '*
    '**********
    p借入金マスタ.据置回数 = DateDiff("m", w実行支払年月, p借入金マスタ.初回返済年月)
    w据置回数 = p借入金マスタ.据置回数
    p借入金マスタ.支払回数 = DateDiff("m", p借入金マスタ.初回返済年月, _
                                           p借入金マスタ.最終返済年月) + 1
                                           
    sv支払回数 = p借入金マスタ.支払回数             ' 08/07/19 V188  解約無い時の支払回数
    
                                           
    '***V188にて INPUT時　で入力する  08/07/07
    'wd01 = MXA030_翌営業年月日計算(CDate(p借入金マスタ.最終返済年月), p借入金マスタ.支払日, p借入金マスタ.営業日区分) ' 07/01/30 V180
    'p借入金マスタ.最終返済実行日 = Format(wd01, "yyyy/mm/dd") 'V180 07/01/21
                                                           
    If Not IsNull(w解約実行日) Then
    
    'MBD010_実行解約年月(p年月日 As Variant, _
                                    p初回返済年月 As Variant, p初回返済実行日 As Variant, _
                                    p最終返済年月 As Variant, p最終返済実行日 As Variant, _
                            p営業日区分 As Integer, p判断 As String) As Variant ' 07/01/30 V180
'
    
    
    
        w実行解約年月 = MBD010_実行解約年月(w解約実行日, _
                               p借入金マスタ.初回返済年月, p借入金マスタ.初回返済実行日, _
                               p借入金マスタ.最終返済年月, p借入金マスタ.最終返済実行日, _
                               p借入金マスタ.支払日, p借入金マスタ.営業日区分)      ' 08/07/24 V188
        'If Format(w実行解約年月, "yyyy/mm/dd") > Format(p借入金マスタ.最終返済年月, "yyyy/mm/dd") Then  ' 08/07/20 V188
        '    w実行解約年月 = p借入金マスタ.最終返済年月                  ' 08/07/20 V188
        'End If
        ' 08/07/20 V188
        
        '*** 実行日の締年月算出 10/01/16
        'w実行支払年月 = Format(p借入金マスタ.実行日, "yyyy/mm/01")  '10/01/16
        'w実行支払年月 = MXA030_翌営業年月日計算(w実行支払年月, p借入金マスタ.支払日, p借入金マスタ.営業日区分)
        'If Format(w実行支払年月, "yyyy/mm/dd") <= Format(p借入金マスタ.実行日, "yyyy/mm/dd") Then '10/01/16
        '    w実行支払年月 = DateAdd("m", 1, w実行支払年月)                                          '10/01/16
        'End If                  '10/01/16
        
        'w実行支払年月 = MXA030_実行支払年月(p借入金マスタ.実行日, w支払日, p借入金マスタ.営業日区分, "*")     ' 07/01/30 V180
        '据置期間の解約
        If w解約実行日 <= p借入金マスタ.初回返済実行日 Then
        '    '実行解約年月の調整
        '    wyy = Year(w解約実行日)                        ' 08/07/20 V188
        '    wmm = Month(w解約実行日) - 1                   ' 08/07/20 V188
        '    w調整解約年月 = CDate(CStr(wyy) + "/" + CStr(wmm) + "/" + CStr(1))  ' 08/07/20 V188
        '    w調整解約年月 = MXA030_翌営業年月日計算(w調整解約年月, p借入金マスタ.支払日, p借入金マスタ.営業日区分) ' 07/01/30 V180
        '    If Format(w調整解約年月, "yyyy/mm/dd") = Format(w解約実行日, "yyyy/mm/dd") Then ' 08/07/20 V188
        '        w実行解約年月 = CDate(CStr(wyy) + "/" + CStr(wmm) + "/" + CStr(1))  ' 08/07/20 V188
        '    End If
            w据置回数 = DateDiff("m", w実行支払年月, w実行解約年月) + 1
            p借入金マスタ.支払回数 = 0
        '据置期間後の解約
        Else
           
            p借入金マスタ.支払回数 = DateDiff("m", p借入金マスタ.初回返済年月, w実行解約年月) + 1
            '解約年月=初回支払年月かつ手入力初回返済実行日が標準初回返済実行日よ以前の時　支払回数の調整
            'If Format(w解約年月, "yyyy/mm/dd") = Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd") Then ' 08/07/20 V188
            '    GDate1 = MBD010_利息計算年月日(p借入金マスタ.初回返済年月, p借入金マスタ.支払日, _
            '        p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)  ' 08/07/20 V188
            '    If Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") < _
            '        Format(GDate1, "yyyy/mm/dd") Then                               ' 08/07/20 V188
            '        p借入金マスタ.支払回数 = p借入金マスタ.支払回数 + 1             ' 08/07/20 V188
            '    End If                                                              ' 08/07/20 V188
            'End If                                                                  ' 08/07/20 V188
            '解約年月=最終支払年月かつ手入力最終返済実行日が標準初回返済実行日よ以降の時　支払回数の調整
            'If Format(w解約年月, "yyyy/mm/dd") = Format(p借入金マスタ.最終返済年月, "yyyy/mm/dd") Then ' 08/07/20 V188
            '    GDate1 = MBD010_利息計算年月日(p借入金マスタ.最終返済年月, p借入金マスタ.支払日, _
            '         p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)  ' 08/07/20 V188
            '    If Format(p借入金マスタ.最終返済実行日, "yyyy/mm/dd") < _
            '        Format(GDate1, "yyyy/mm/dd") Then                               ' 08/07/20 V188
            '        p借入金マスタ.支払回数 = p借入金マスタ.支払回数 - 1             ' 08/07/20 V188
            '    End If                                                              ' 08/07/20 V188
            'End If                                                                  ' 08/07/20 V188
        End If
        'w年月日 = DateAdd("m", -1, w解約実行日)
        'w年月日 = C年月日.GetDate("月末", w年月日)
        'w銀行マスタ = MAA030_銀行マスタRead(p借入金マスタ.銀行番号)
        'If p借入金マスタ.支払日  = 31 Then
        '    w年月日 = MXA030_翌営業年月日計算(w年月日, 31)
        '    If w年月日 = w解約実行日 Then
        '        p借入金マスタ.支払回数 = p借入金マスタ.支払回数 - 1
        '    End If
        'End If
        
            
    End If
    
    
    '解約の時　支払回数　据置回数　異常を正常に戻す
    If w据置回数 > p借入金マスタ.据置回数 Then      ' 08/07/21 V188
        w据置回数 = p借入金マスタ.据置回数          ' 08/07/21 V188
        p借入金マスタ.支払回数 = 1                  ' 08/07/21 V188
    End If                                          ' 08/07/21 V188
    
    If p借入金マスタ.支払回数 > sv支払回数 Then     ' 08/07/21 V188
        p借入金マスタ.支払回数 = sv支払回数         ' 08/07/21 V188
    End If                                          ' 08/07/21 V188
    
     
    
    
    ' =========================================
    '         グローバルテーブル　クリア
    ' =========================================
    If pClear Then
        ReDim G借入金テーブル(0)
    End If
    
    'If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then          ' 2009/12/20
    '    wGAP = Fix((DateDiff("m", w実行支払年月, p借入金マスタ.初回返済年月) - 2) / p借入金マスタ.返済単位月数) '09/12/20 V185
    '    wGAP = wGAP * p借入金マスタ.返済単位月数                        '08/03/14 V185
    '    w一括支払利息開始年月 = DateAdd("m", -wGAP, p借入金マスタ.初回返済年月) '08/03/14 V185
    'Else                                                                ' 2009/12/20
    
        If p借入金マスタ.実行日 = p借入金マスタ.初回返済実行日 Then     '10/05/06 V195
            If p借入金マスタ.利息支払方法 = 1 Then                          '10/05/06 V195
                p借入金マスタ.金利初回年月 = DateAdd("M", p借入金マスタ.返済単位月数 - 1, w実行支払年月) '10/05/06 V195
            Else                                                            '10/05/06 V195
                p借入金マスタ.金利初回年月 = w実行支払年月                  '10/05/06 V195
            End If                                                          '10/05/06 V195
        Else                                                                '10/05/06 V195
            p借入金マスタ.金利初回年月 = Format(p借入金マスタ.金利初回年月, "yyyy/mm/01") '09/12/22
        End If                                                              '10/05/06 V195
        
        
        '***金利初回年月<=w実行支払年月 10/01/17
        If Format(p借入金マスタ.金利初回年月, "yyyy/mm/dd") <= Format(w実行支払年月, "yyyy/mm/dd") Then '10/01/17
            p借入金マスタ.金利初回年月 = w実行支払年月                      '10/01/17
        End If                                                              '10/01/17
        
        If p借入金マスタ.手入力区分 = P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
            w一括支払利息開始年月 = p借入金マスタ.金利初回年月              ' 2009/12/22
        Else
            w一括支払利息開始年月 = p借入金マスタ.初回返済年月              ' 2009/12/22
        End If

    'End If                                                              ' 2009/12/22
    
    
    ' *****************************************
    ' *
    ' *          借入金テーブル セット
    ' *
    ' *****************************************
    ' -----------------------------------------
    '            実行日のＤＡＴＡ
    ' -----------------------------------------
    
  If p借入金マスタ.実行日 <> p借入金マスタ.初回返済実行日 Then  '10/05/06 V195
    
    w内入開始年月日 = Null                   ' 08/12/05 V189
    w内入終了年月日 = Null                   ' 08/12/05 V189
    w融資残高 = p借入金マスタ.融資金額       ' 08/12/05 V189
    
    
    wTable.借入番号 = p借入金マスタ.借入番号
    wTable.返済回数 = 0
    wTable.据置x回目 = 2                   ' 08/12/06 V189 借入明細出力順で使用
    wTable.返済予定年月 = C年月日.GetDate("月始", p借入金マスタ.実行日)
    
    
    
    wTable = MBD010_借入金テーブルRead(wTable)
        
        wTable.実際年月日 = p借入金マスタ.実行日
        OLD実際年月日 = wTable.実際年月日           ' 07/03/24 V180
        NEW実際年月日 = wTable.実際年月日           ' 07/03/24 V180
        wTable.元金額 = 0
        wTable.保証料 = 0
        wTable.融資残高 = p借入金マスタ.融資金額
        wTable.利率 = p借入金マスタ.利率
        
        '***** 返済単位月数＝１　は　利息支払方法　０（毎月）をセット
        If p借入金マスタ.返済単位月数 = 1 Then  '2009/12/20
            p借入金マスタ.利息支払方法 = 0      '2009/12/20
        End If                                  '2009/12/20
        
    
        If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
          '利息先払
          If p借入金マスタ.利息支払方法 = 0 Then                                    ' 07 01/31 V180
            '利息毎月支払
            GDate1 = MBD010_利息計算年月日(p借入金マスタ.金利初回年月, p借入金マスタ.支払日, _
                     p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)  ' 07/01/30 V180
                     
            If Format(p借入金マスタ.金利初回年月, "yyyy/mm/dd") = Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd") Then   '10/01/01
                If Format(GDate利息対象年月日, "yyyy/mm/dd") <> Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then  ' 08/12/21 V189
                    GDate1 = p借入金マスタ.初回返済実行日                       ' 08/07/19 V188
                    GDate利息対象年月日 = p借入金マスタ.初回返済実行日          ' 08/07/19 V188
                End If                                                          ' 08/07/22 V188
            End If                                                          ' 08/07/19 V188
            
            wTable.日割日数 = DateDiff("d", p借入金マスタ.実行日, GDate1) + 1
            w内入開始年月日 = p借入金マスタ.実行日          ' 08/12/05 V189
            w内入終了年月日 = GDate1                        ' 08/12/05 V189
            
            
            GDate1利息対象年月日 = GDate利息対象年月日      'V182 2008/02/28
            wTable.利息対象期間日数 = wTable.日割日数      '09/12/29
            w利息対象内入開始年月日 = p借入金マスタ.実行日          ' 08/12/08 V189
            w利息対象内入終了年月日 = GDate1利息対象年月日          ' 08/12/08 V189
            
            If p借入金マスタ.利息控除区分 = 1 Or p借入金マスタ.利息控除区分 = 3 Then  ' 07/03/05 V180
                wTable.日割日数 = wTable.日割日数 - 1                                 ' 07/03/05 V180
            End If                                                                    ' 07/03/05 V180
            
            If Format(p借入金マスタ.金利初回年月, "yyyy/mm/dd") = _
               Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd") _
               And Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd") = _
                   Format(p借入金マスタ.最終返済年月, "yyyy/mm/dd") _
               And (p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3) Then  '09/12/31
                    wTable.日割日数 = wTable.日割日数 - 1               '09/12/31
            End If                                                  '09/12/31

            
            'If p借入金マスタ.金利計算年間日数 = 0 Then                                               ' 07/01/31 V180
            '    wTable.利息額 = Fix(p借入金マスタ.融資金額 * CCur(p借入金マスタ.利率) * wTable.日割日数 / 36500)
            'Else                                                                    ' 07/01/31 V180
            '    wTable.利息額 = Fix(p借入金マスタ.融資金額 * CCur(p借入金マスタ.利率) * wTable.日割日数 / 36000) ' 07/01/31 V180
            'End If                                                                  ' 07/01/31 V180
            wTable.利息額 = MBD010_利息計算小数点5桁(p借入金マスタ.利率, p借入金マスタ.融資金額, _
                                    wTable.日割日数, p借入金マスタ.金利計算年間日数) '09/12/30
                                    
            
            
          Else                                                                      ' 07/01/31 V180
            '利息一括支払
            'If p借入金マスタ.初回返済年月 = p借入金マスタ.最終返済年月 Then         ' 07/01/31 V180
            '    GDate1 = MBD010_利息計算年月日(p借入金マスタ.最終返済年月, p借入金マスタ.支払日, _
            '         p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)  ' 07/01/30 V180
            'Else                                                                    ' 07/01/31 V180
            '    GDate1 = MBD010_利息計算年月日(p借入金マスタ.初回返済年月, p借入金マスタ.支払日, _
            '         p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)  ' 07/01/30 V180
            'End If                                                                  ' 07/01/31 V180
            
            GDate1 = MBD010_利息計算年月日(w一括支払利息開始年月, p借入金マスタ.支払日, _
                                           p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分) '09/12/20 V185
                                           
            If Format(p借入金マスタ.金利初回年月, "yyyy/mm/dd") = Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd") Then   '10/01/01
                If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                   Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then  '10/01/01
                    GDate1 = p借入金マスタ.初回返済実行日                   '10/01/01
                    GDate利息対象年月日 = p借入金マスタ.初回返済実行日      '10/01/01
                End If                                                      '10/01/01
            End If                                                          '10/01/01
                                           
                                           
            wTable.日割日数 = DateDiff("d", p借入金マスタ.実行日, GDate1) + 1   ' 07/01/31 V180
            w内入開始年月日 = p借入金マスタ.実行日          ' 08/12/05 V189
            w内入終了年月日 = GDate1                        ' 08/12/05 V189
            
            GDate1利息対象年月日 = GDate利息対象年月日      'V182 2008/02/28
            wTable.利息対象期間日数 = wTable.日割日数       '09/12/29
            w利息対象内入開始年月日 = p借入金マスタ.実行日          ' 08/12/08 V189
            w利息対象内入終了年月日 = GDate1利息対象年月日          ' 08/12/08 V189
            
            If p借入金マスタ.利息控除区分 = 1 Or p借入金マスタ.利息控除区分 = 3 Then  ' 07/03/05 V180
                wTable.日割日数 = wTable.日割日数 - 1                                 ' 07/03/05 V180
            End If                                                                    ' 07/03/05 V180
            
            If Format(p借入金マスタ.金利初回年月, "yyyy/mm/dd") = _
               Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd") _
               And Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd") = _
                   Format(p借入金マスタ.最終返済年月, "yyyy/mm/dd") _
               And (p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3) Then  '09/12/31
                    wTable.日割日数 = wTable.日割日数 - 1               '09/12/31
            End If                                                  '09/12/31

            
            
            'If (p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3) _
            '    And p借入金マスタ.初回返済年月 = p借入金マスタ.最終返済年月 Then      ' 07/03/05 V180
            '    wTable.日割日数 = wTable.日割日数 - 1                                 ' 07/03/05 V180
            'End If                                                                    ' 07/03/05 V180
          
          
          End If                                                                    ' 07/03/05 V180
            
          
          'If p借入金マスタ.金利計算年間日数 = 0 Then                            'V180 07/02/01
          '  wTable.利息額 = Fix(p借入金マスタ.融資金額 * CCur(p借入金マスタ.利率) * wTable.日割日数 / 36500)
          'Else                                                                      'V180 07/02/01
          '  wTable.利息額 = Fix(p借入金マスタ.融資金額 * CCur(p借入金マスタ.利率) * wTable.日割日数 / 36000)
          'End If                                                                    'V180 07/02/01
          wTable.利息額 = MBD010_利息計算小数点5桁(p借入金マスタ.利率, p借入金マスタ.融資金額, _
                                    wTable.日割日数, p借入金マスタ.金利計算年間日数) '09/12/30
            
          
           
          If wTable.日割日数 = 1 And (p借入金マスタ.利息控除区分 = 0 Or _
                                       p借入金マスタ.利息控除区分 = 2) Then         ' 07/02/22 V180
            wp日数 = 1                         ' 07/02/21 V180
            wTable.日割日数 = 0                ' 07/02/21 V180
            wTable.利息対象期間日数 = 0         'V182 2008/01/28
            wTable.利息額 = 0                  ' 07/02/21 V180
            wTable.返済金額 = 0                ' 07/02/21 V180
          End If
             
        Else
          '利息後払
            wTable.日割日数 = 0
            wTable.利息対象期間日数 = 0         'V182 2008/01/28
            wTable.利息額 = 0
            wTable.返済金額 = 0
        End If
        
        wTable.返済金額 = wTable.利息額                                             'V18007/02/01
        
    If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息後払") Then           ' 08/12/06 V189
    
        Call MBD010_内入処理(p借入金マスタ)         ' 08/12/05 V189
    End If                                                                          ' 08/12/06 V189
        
    wTable.据置x回目 = 2                    ' 09/01/16 V189
    wTable.利息計算年月日 = wTable.実際年月日   '10/01/04
    Call MBD010_借入金テーブルWrite(wTable, p借入金マスタ)       '1016/09/23
    
    
    If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then           ' 08/12/06 V189
    
        Call MBD010_内入処理(p借入金マスタ)         ' 08/12/05 V189
    End If                                                                          ' 08/12/06 V189
    
         
    ' -----------------------------------------
    '            据置期間のＤＡＴＡ
    ' -----------------------------------------
    For wRecCount = 0 To w据置回数 - 1
        
        'w内入開始年月日 = Null                  ' 08/12/05 V189
        'w内入終了年月日 = Null                  ' 08/12/05 V189
        
        
        w返済予定年月 = DateAdd("m", wRecCount, w実行支払年月)
        w実際年月日 = MXA030_翌営業年月日計算(w返済予定年月, p借入金マスタ.支払日, p借入金マスタ.営業日区分) ' 07/01/30 V180
        
        
        
        
        
        w内入開始年月日 = Null                                          '10/02/07
        w内入終了年月日 = Null                                          '10/02/07
        
        '***元金、利息がゼロの時　w内入開始年月日　w内入終了年月日 の初期値セット
     'If Format(w返済予定年月, "yyyy/mm/dd") >= Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd") Then
     'Else                                                               '10/02/07
     
     '   If Format(p借入金マスタ.金利初回年月, "yyyy/mm/dd") < Format(w返済予定年月, "yyyy/mm/dd") Then
     '       wW = DateDiff("M", w実行支払年月, w返済予定年月)    '10/02/05
     '       w返済単位判定 = wRecCount + p借入金マスタ.返済単位月数 - wW + 1 '10/02/05
     '       w返済単位回数 = Fix(w返済単位判定 / p借入金マスタ.返済単位月数) '10/02/05
     '       wx = w返済単位判定 - w返済単位回数 * p借入金マスタ.返済単位月数 '10/02/05 後払
     '       wY = p借入金マスタ.返済単位月数                                 '10/02/05
     '       wZ = wY - wx                                                    '10/02/05 先払
     '
     '       If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then       '10/02/05
     '           '*利息先払
     '           GDate1 = DateAdd("m", wZ, w返済予定年月)                                 '10/02/05
     '           If Format(GDate1, "yyyy/mm/dd") >= Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd") Then '10/02/07
     '               GDate1 = p借入金マスタ.初回返済年月                     '10/02/07
     '           End If                                                      '10/02/07
     '           GDate1 = MBD010_利息計算年月日(GDate1, p借入金マスタ.支払日, _
     '                p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)      '10/02/05
     '           GDate2 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
     '                p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)      '10/02/05
     '       Else                                                                        '10/02/05
     '           '*利息後払
     '           GDate1 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
     '                p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)      '10/02/05
     '           GDate2 = DateAdd("m", -wx, w返済予定年月)                                '10/02/05
     '           GDate2 = MBD010_利息計算年月日(GDate2, p借入金マスタ.支払日, _
     '                p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)      '10/02/05
     '       End If                                                                      '10/02/05
     '
     '   Else                                                                '10/02/05
     '       If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then '10/02/05
     '            GDate2 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
     '                p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)      '10/02/05
     '            GDate1 = MBD010_利息計算年月日(p借入金マスタ.金利初回年月, p借入金マスタ.支払日, _
     '                p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)      '10/02/05
     '       Else                                                            '10/02/05
     '           GDate2 = p借入金マスタ.実行日                               '10/02/05
     '           GDate1 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
     '                p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)      '10/02/05
     '       End If                                                          '10/02/05
     '   End If                                                              '10/02/05
     '
     '   w内入開始年月日 = GDate2                                                    '10/02/05
     '   w内入終了年月日 = GDate1                                                    '10/02/05
     '
     'End If                                                                 '10/02/07
        
        wTable.借入番号 = p借入金マスタ.借入番号
        wTable.返済回数 = 0
        wTable.据置x回目 = 2                   ' 08/12/06 V189
        wTable.返済予定年月 = w返済予定年月
        wTable.実際年月日 = w実際年月日
        OLD実際年月日 = NEW実際年月日               ' 07/03/24 V180
        NEW実際年月日 = wTable.実際年月日           ' 07/03/24 V180
        
        wTable = MBD010_借入金テーブルRead(wTable)
                
            wTable.実際年月日 = w実際年月日
            wTable.元金額 = 0
            wTable.保証料 = 0
            wTable.融資残高 = w融資残高             ' 08/12/05 V189
            wTable.利率 = MBD010_金利参照(p借入金マスタ, w返済予定年月)
               
            '--------------
            '** 日割日数 **
            '--------------
            If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
              '利息先払い
              If p借入金マスタ.利息支払方法 = 0 Then                                    ' 07 01/31 V180
                '利息毎月支払
                GDate1 = DateAdd("m", 1, w返済予定年月)
                GDate1 = MBD010_利息計算年月日(GDate1, p借入金マスタ.支払日, _
                     p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)  ' 07/01/30 V180
                
                     
                If wRecCount = w据置回数 - 1 Then                               ' 08/07/19 V188
                    If Format(GDate利息対象年月日, "yyyy/mm/dd") <> Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then  ' 08/07/22 V188
                        GDate1 = p借入金マスタ.初回返済実行日                   ' 08/07/19 V188
                        GDate利息対象年月日 = p借入金マスタ.初回返済実行日      ' 08/07/19 V188
                    End If                                                      ' 08/07/22 V188
                End If                                                          ' 08/07/19 V188
                 
                     
                GDate1利息対象年月日 = GDate利息対象年月日      'V182 2008/02/28
                 
                GDate2 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
                     p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)  ' 07/01/30 V180
                wTable.日割日数 = DateDiff("d", GDate2, GDate1)                     ' 07/01/31 V180
                w内入開始年月日 = GDate2                    ' 08/12/05 V189
                w内入終了年月日 = GDate1                    ' 08/12/05 V189
                
                GDate2利息対象年月日 = GDate利息対象年月日      'V182 2008/02/28
                wTable.利息対象期間日数 = wTable.日割日数       '09/12/29
                w利息対象内入開始年月日 = GDate2利息対象年月日          ' 08/12/08 V189
                w利息対象内入終了年月日 = GDate1利息対象年月日          ' 08/12/08 V189
                wTable.日割日数 = wTable.日割日数 + wp日数
                wTable.利息対象期間日数 = wTable.日割日数       '09/12/29
                wp日数 = 0
                
                GDate1 = DateAdd("m", 1, w返済予定年月)                         'V180 07/02/01
                If GDate1 = p借入金マスタ.最終返済年月 Then                     'V180 07/02/01
                    If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then
                        wTable.日割日数 = wTable.日割日数 - 1                       'V180 07/02/01
                    End If
                End If                                                          'V180 07/02/01
                    
                '***据え置き期間　実行日に金利初回年月まで徴収した調整
                If Format(w返済予定年月, "yyyy/mm/dd") < Format(p借入金マスタ.金利初回年月, "yyyy/mm/dd") Then  ' 09/12/22
                    wTable.日割日数 = 0             ' 09/12/22
                    wTable.利息対象期間日数 = 0     ' 09/12/22
                    w内入開始年月日 = Null          '10/02/13
                    w内入終了年月日 = Null          '10/02/13
                End If                              ' 09/12/22
                
                    
              Else                                                                  ' 07/01/31 V180
                '利息一括支払
                w有効F = MBD010_借入金据置期間一括利息支払(wTable.返済予定年月, w一括支払利息開始年月 _
                         , p借入金マスタ.初回返済年月, p借入金マスタ.利息区分, p借入金マスタ.返済単位月数) '08/03/14 V185
                If w有効F = 1 Then                                      '08/03/14 V185
                    GDate1 = DateAdd("m", p借入金マスタ.返済単位月数, w返済予定年月) '08/03/14 V185
                    GDate2 = GDate1                     '10/02/09
                    GDate1 = MBD010_利息計算年月日(GDate1, p借入金マスタ.支払日, _
                     p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)  ' 08/03/14 V185
                     
                    'If wRecCount = w据置回数 - p借入金マスタ.返済単位月数 _
                    '    And Format(GDate利息対象年月日, "yyyy/mm/dd") _
                    '        <> Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then ' 10/02/06
                    If Format(GDate2, "yyyy/mm/dd") = Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd") _
                        And Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") _
                         <> Format(GDate利息対象年月日, "yyyy/mm/dd") Then          '10/02/09
                        GDate1 = p借入金マスタ.初回返済実行日                       ' 10/01/01
                        GDate利息対象年月日 = p借入金マスタ.初回返済実行日          ' 10/01/01
                    End If                                                          ' 10/01/01
                     
                    GDate1利息対象年月日 = GDate利息対象年月日      '08/03/14 V185
                    
                    GDate2 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
                     p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)  ' 08/03/14 V185
                    wTable.日割日数 = DateDiff("d", GDate2, GDate1)             ' 08/03/14 V185
                    w内入開始年月日 = GDate2                    ' 08/12/05 V189
                    w内入終了年月日 = GDate1                    ' 08/12/05 V189
                    
                    GDate2利息対象年月日 = GDate利息対象年月日      '08/03/14 V185
                    wTable.利息対象期間日数 = wTable.日割日数       '09/12/29
                    w利息対象内入開始年月日 = GDate2利息対象年月日          ' 08/12/08 V189
                    w利息対象内入終了年月日 = GDate1利息対象年月日          ' 08/12/08 V189
                    
                    GDate1 = DateAdd("m", p借入金マスタ.返済単位月数, w返済予定年月)    '10/01/01
                    If GDate1 = p借入金マスタ.最終返済年月 Then                         '10/01/01
                        If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then    '10/01/01
                            wTable.日割日数 = wTable.日割日数 - 1                       '10/01/01
                        End If                                                          '10/01/01
                    End If                                                              '10/01/01
                    
                Else                                        '08/03/14 V185
                    wTable.日割日数 = 0                                                 ' 07/01/31 V180
                    wTable.利息対象期間日数 = 0                                     'V182 2008/01/28
                End If                                      '08/03/14 V185
                
                'wTable.日割日数 = 0                         ' 09/12/20 V188
                'wTable.利息対象期間日数 = 0                 ' 09/12/20 V188
                
                
                
              End If                                                                ' 07/01/31 V180
                
            Else
              '利息後払い
              If p借入金マスタ.利息支払方法 = 0 Then                                ' 07/01/31 V180
                '利息毎月支払
                If w返済予定年月 = p借入金マスタ.金利初回年月 Then                  ' 07/01/31 V180
                    GDate1 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
                     p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)  ' 07/01/30 V180
                     
                    'If wRecCount = w据置回数 - 1 Then                               ' 08/07/19 V188
                    '    GDate1 = p借入金マスタ.初回返済実行日                       ' 08/07/19 V188
                    '    GDate利息対象年月日 = p借入金マスタ.初回返済実行日          ' 08/07/19 V188
                    'End If                                                          ' 08/07/19 V188
                     
                    wTable.日割日数 = DateDiff("d", p借入金マスタ.実行日, GDate1) + 1   ' 07/01/31 V180
                    w内入開始年月日 = p借入金マスタ.実行日      ' 08/12/05 V189
                    w内入終了年月日 = GDate1                    ' 08/12/05 V189
                    
                    GDate1利息対象年月日 = GDate利息対象年月日      'V182 2008/02/28
                    wTable.利息対象期間日数 = wTable.日割日数       '09/12/29
                    w利息対象内入開始年月日 = p借入金マスタ.実行日          ' 08/12/08 V189
                    w利息対象内入終了年月日 = GDate1利息対象年月日          ' 08/12/08 V189
                
                    If p借入金マスタ.利息控除区分 = 1 Or p借入金マスタ.利息控除区分 = 3 Then ' 07/01/31 V180
                        wTable.日割日数 = wTable.日割日数 - 1                       ' 07/01/31 V180
                    End If                                                          ' 07/01/31 V180
                Else                                                                ' 07/01/31 V180
                    If w返済予定年月 > p借入金マスタ.金利初回年月 Then              ' 07/01/31 V180
                        GDate1 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
                          p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分) ' 07/01/30 V180
                          
                        'If wRecCount = w据置回数 - 1 Then                           ' 08/07/19 V188
                        '    GDate1 = p借入金マスタ.初回返済実行日                   ' 08/07/19 V188
                        '    GDate利息対象年月日 = p借入金マスタ.初回返済実行日      ' 08/07/19 V188
                        'End If                                                      ' 08/07/19 V188
                          
                        GDate1利息対象年月日 = GDate利息対象年月日      'V182 2008/02/28
                        
                        GDate2 = DateAdd("m", -1, w返済予定年月)                    ' 07/01/31 V180
                        GDate2 = MBD010_利息計算年月日(GDate2, p借入金マスタ.支払日, _
                          p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分) '07/01/30 V180
                        wTable.日割日数 = DateDiff("d", GDate2, GDate1)              ' 07/01/31 V180
                        w内入開始年月日 = GDate2                    ' 08/12/05 V189
                        w内入終了年月日 = GDate1                    ' 08/12/05 V189
                        
                        
                        GDate2利息対象年月日 = GDate利息対象年月日      'V182 2008/02/28
                        wTable.利息対象期間日数 = wTable.日割日数       '09/12/29
                        w利息対象内入開始年月日 = GDate2利息対象年月日          ' 08/12/08 V189
                        w利息対象内入終了年月日 = GDate1利息対象年月日          ' 08/12/08 V189
                    Else                                                            ' 07/01/31 V180
                    
                              
                        wTable.日割日数 = 0                                         ' 07/01/31 V180
                        wTable.利息対象期間日数 = 0         'V182 2008/01/29
                    End If                                                          ' 07/01/31 V180
                    
                End If                                                              ' 07/01/31 V180
                
              Else                                                                  ' 07/01/31 V180
                '利息一括支払
                w有効F = MBD010_借入金据置期間一括利息支払(wTable.返済予定年月, w一括支払利息開始年月 _
                         , p借入金マスタ.初回返済年月, p借入金マスタ.利息区分, p借入金マスタ.返済単位月数) '08/03/14 V185
                If w有効F = 1 Then                                      '08/03/14 V185
                    If Format(w返済予定年月, "yyyy/mm/dd") = Format(w一括支払利息開始年月, "yyyy/mm/dd") Then ' 2009/12/20
                        GDate1 = p借入金マスタ.実行日           ' 2009/12/20
                        GDate利息対象年月日 = p借入金マスタ.実行日  ' 2009/12/20
                    Else                                        ' 2009/12/20
                        GDate1 = DateAdd("m", -p借入金マスタ.返済単位月数, w返済予定年月) '08/03/14 V185
                    
                        GDate1 = MBD010_利息計算年月日(GDate1, p借入金マスタ.支払日, _
                            p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)  ' 08/03/14 V185
                     
                    End If                                          ' 2009/12/20
                     
                     
                    GDate1利息対象年月日 = GDate利息対象年月日      '08/03/14 V185
                    
                    w初回利息F = 0                                  '08/03/14 V185
                    
                    If GDate1 < p借入金マスタ.実行日 Then           '08/03/14 V185
                        GDate1 = DateAdd("d", -1, p借入金マスタ.実行日)             '08/03/14 V185
                        GDate1利息対象年月日 = DateAdd("d", -1, p借入金マスタ.実行日) '08/03/14 V185
                        w初回利息F = 1                              '08/03/04 V185
                    Else
                        If GDate1 = p借入金マスタ.実行日 Then       '09/12/31
                            w初回利息F = 1                          '09/12/31
                        End If                                      '09/12/31
                    End If
                                                                    '08/03/04 V185
                                                                    
                    GDate2 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
                     p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)  ' 08/03/14 V185
                     
                    If wRecCount = w据置回数 - 1 Then                                   ' 08/07/19 V188
                        If Format(GDate利息対象年月日, "yyyy/mm/dd") <> Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then ' 08/07/19 V188
                            GDate2 = p借入金マスタ.初回返済実行日                       ' 08/07/19 V188
                            GDate利息対象年月日 = p借入金マスタ.初回返済実行日         ' 08/07/19 V188
                        End If                                                          ' 08/07/19 V188
                    End If                                                              ' 08/07/19 V188
                    
                    If Format(w返済予定年月, "yyyy/mm/dd") = Format(w一括支払利息開始年月, "yyyy/mm/dd") Then ' 2009/12/20
                        wTable.日割日数 = DateDiff("d", GDate1, GDate2) + 1 ' 2009/12/20
                    Else                                                    ' 2009/12/20
                        wTable.日割日数 = DateDiff("d", GDate1, GDate2)             ' 08/03/14 V185
                    End If                                                  ' 2009/12/20
                    
                    w内入開始年月日 = GDate1                    ' 08/12/05 V189
                    w内入終了年月日 = GDate2                    ' 08/12/05 V189
                    
                    GDate2利息対象年月日 = GDate利息対象年月日      '08/03/14 V185
                    wTable.利息対象期間日数 = wTable.日割日数       '09/12/29
                    w利息対象内入開始年月日 = GDate1利息対象年月日          ' 08/12/08 V189
                    w利息対象内入終了年月日 = GDate2利息対象年月日          ' 08/12/08 V189
                    If w初回利息F = 1 And (p借入金マスタ.利息控除区分 = 1 Or p借入金マスタ.利息控除区分 = 3) Then ' 08/03/141 V185
                        wTable.日割日数 = wTable.日割日数 - 1                       ' 07/01/31 V180
                    End If                                                          ' 07/01/31 V180
                Else                                        '08/03/14 V185
                    wTable.日割日数 = 0                                                 ' 07/01/31 V180
                    wTable.利息対象期間日数 = 0                                     'V182 2008/01/28
                End If                                      '08/03/14 V185
                
                'wTable.日割日数 = 0                         ' 09/12/20 V188
                'wTable.利息対象期間日数 = 0                  ' 09/12/20 V188
                
              End If                                                                ' 07/01/31 V180
              
            End If                                                                  ' 07/01/31 V180
            
              
            '** 解約 **
            If Not IsNull(w解約実行日) Then
                If Format(w実際年月日, "yyyy/mm/dd") >= Format(w解約実行日, "yyyy/mm/dd") Then
                     
                    If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
                      '利息先払い
                      If p借入金マスタ.利息支払方法 = 0 Then                        ' 07/01/31 V180
                        '利息毎月支払
                        
                        '実行日に金利初回まで徴収した調整
                        If Format(p借入金マスタ.金利初回年月, "yyyy/mm/dd") _
                                    = Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd") _
                                And Format(p借入金マスタ.金利初回年月, "yyyy/mm/dd") _
                                    = Format(p借入金マスタ.最終返済年月, "yyyy/mm/dd") Then     '10/02/16
                            GDate1 = MBD010_利息計算年月日(p借入金マスタ.金利初回年月, p借入金マスタ.支払日, _
                                p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)   '10/02/16
                        Else                                                                '10/02/16
                                                
                            If Format(w返済予定年月, "yyyy/mm/dd") < Format(p借入金マスタ.金利初回年月, "yyyy/mm/dd") Then ' 09/12/22
                                GDate1 = MBD010_利息計算年月日(p借入金マスタ.金利初回年月, p借入金マスタ.支払日, _
                                    p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)  ' 09/12/22
                            Else                                                            ' 09/12/22
                
                                GDate1 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
                                    p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)  ' 07/01/30 V180
                            End If                                                          ' 09/12/22
                        End If                                                              '10/02/16
                        
                        '*初回返済日が手打ちの時の調整
                        If Format(w返済予定年月, "yyyy/mm/dd") = Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd") _
                           Or (Format(p借入金マスタ.金利初回年月, "yyyy/mm/dd") _
                                    = Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd") _
                                And Format(p借入金マスタ.金利初回年月, "yyyy/mm/dd") _
                                    = Format(p借入金マスタ.最終返済年月, "yyyy/mm/dd")) _
                            And Format(GDate利息対象年月日, "yyyy/mm/dd") _
                                <> Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then  '10/02/16
                            GDate1 = p借入金マスタ.初回返済実行日           '10/02/09
                            GDate利息対象年月日 = p借入金マスタ.初回返済実行日   '10/02/09
                        End If                                              '10/02/09
                        
                        
                        
                            
                        wTable.日割日数 = DateDiff("d", w解約実行日, GDate1) * -1   ' 07/01/31 V180
                        
                        w内入開始年月日 = w解約実行日           '10/01/30
                        w内入終了年月日 = GDate1                '10/01/30
                         
                        GDate1利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                        wTable.利息対象期間日数 = wTable.日割日数       '09/12/29
                        
                        '***金利初回年月=初回返済年月=最終返済年月の時　10/02/13
                        'If Format(MBD010_利息計算年月日(p借入金マスタ.金利初回年月, p借入金マスタ.支払日 _
                        '               , p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分), "yyyy/mm/dd") _
                        '            >= Format(w解約実行日, "yyyy/mm/dd")
                        If Format(p借入金マスタ.金利初回年月, "yyyy/mm/dd") _
                                    = Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd") _
                                And Format(p借入金マスタ.金利初回年月, "yyyy/mm/dd") _
                                    = Format(p借入金マスタ.最終返済年月, "yyyy/mm/dd") Then
                                    
                                'And (p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3) _
                                'Then                                    '10/02/13
                                  
                            'wTable.日割日数 = wTable.日割日数 + 1           '10/02/13
                        Else                                              '10/02/13
                        
                            If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then
                                wTable.日割日数 = wTable.日割日数 - 1                   ' 2012/03/21
                            End If                                                      ' 07/03/05 V180
                        End If                                                          '10/02/13
                        
                      Else                                                          ' 07/03/05 V180
                      
                        '利息一括支払い
                        
                        GDate1 = MBD010_借入金解約据置期間一括利息支払日(w解約実行日, _
                                                  p借入金マスタ.実行日, _
                                                  w返済予定年月, _
                                                  w一括支払利息開始年月, _
                                                  p借入金マスタ.初回返済年月, _
                                                  p借入金マスタ.利息区分, _
                                                  p借入金マスタ.返済単位月数, _
                                                  p借入金マスタ.支払日, _
                                                  p借入金マスタ.営業日区分, _
                                                  p借入金マスタ.利息計算日数区分)    '10/02/02
                                                  
                                                  
                        '*初回返済日が手打ちの時の調整
                        GDate2 = MBD010_利息計算年月日(p借入金マスタ.初回返済年月, p借入金マスタ.支払日, _
                            p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)       '10/02/09
                        If Format(GDate1, "yyyy/mm/dd") = Format(GDate2, "yyyy/mm/dd") _
                              And Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                                  Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then '10/02/15
                            GDate1 = p借入金マスタ.初回返済実行日           '10/02/09
                            GDate利息対象年月日 = p借入金マスタ.初回返済実行日  '10/02/09
                        End If                                            '10/02/09
                                                  
                                                  
                        'GDate1 = p借入金マスタ.初回返済実行日                   ' 09/12/20 V188
                        'GDate利息対象年月日 = p借入金マスタ.初回返済実行日      ' 09/12/20 V188
                                                  
                        'GDate1 = MBD010_利息計算年月日(GDate1, p借入金マスタ.支払日, _
                        '  p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分) '09/12/20 V185
                        wTable.日割日数 = DateDiff("d", w解約実行日, GDate1) * -1   ' 09/12/20 V180
                        
                        w内入開始年月日 = w解約実行日           '10/01/30
                        w内入終了年月日 = GDate1                '10/01/30
                        
                          
                        GDate1利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                        wTable.利息対象期間日数 = wTable.日割日数       '09/12/29
                        
                        '*** 金利初回年月=初回返済年月=最終返済年月の時 10/02/13
                        GDate2 = MBD010_利息計算年月日(p借入金マスタ.金利初回年月, p借入金マスタ.支払日 _
                                       , p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)
                        If Format(p借入金マスタ.金利初回年月, "yyyy/mm/dd") = _
                            Format(p借入金マスタ.最終返済年月, "yyyy/mm/dd") Then   '10/02/13
                            If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                                Format(p借入金マスタ.最終返済実行日, "yyyy/mm/dd") Then '10/02/13
                                GDate2 = p借入金マスタ.最終返済実行日                    '10/02/13
                            End If                                                      '10/02/13
                        End If                                                          '10/02/13
                        
                        'If Format(GDate2, "yyyy/mm/dd")
                        '            >= Format(w解約実行日, "yyyy/mm/dd")
                        If Format(p借入金マスタ.金利初回年月, "yyyy/mm/dd") _
                                    = Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd") _
                                And Format(p借入金マスタ.金利初回年月, "yyyy/mm/dd") _
                                    = Format(p借入金マスタ.最終返済年月, "yyyy/mm/dd") Then     '10/02/16
                                'And (p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3) _
                                Then                                    '10/02/13
                                  
                            'wTable.日割日数 = wTable.日割日数 + 1           '10/02/15
                        Else                                              '10/02/13
                        
                            GDate2 = MBD010_利息計算年月日(DateAdd("m", -p借入金マスタ.返済単位月数, p借入金マスタ.最終返済年月), _
                             p借入金マスタ.支払日, p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)    '10/02/11
                        
                            If Format(GDate2, "yyyy/mm/dd") >= _
                                Format(w解約実行日, "yyyy/mm/dd") _
                               Or wTable.日割日数 > 0 Then                     '10/02/11
                                If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then    '08/03/16 V185
                                    wTable.日割日数 = wTable.日割日数 - 1           '08/03/16 V185
                                End If                                              '08/03/16 V185
                            Else                                                    '10/02/07
                                'If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then    '08/03/16 V185
                                '    wTable.日割日数 = wTable.日割日数 + 1           '10/02/07
                                'End If
                            End If                                                  '10/02/07
                        End If                                                    '10/02/13
                          
                        'Else                                                        ' 07/01/31 V180
                        '  GDate1 = MBD010_利息計算年月日(p借入金マスタ.初回返済年月, p借入金マスタ.支払日, _
                        '    p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分) '07/01/30 V180
                        '  wTable.日割日数 = DateDiff("d", w解約実行日, GDate1) * -1 ' 07/01/31 V180
                          
                        '  GDate1利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                        '  wTable.利息対象期間日数 = DateDiff("d", w解約実行日, GDate1利息対象年月日) * -1 'V182 2008/01/29
                          
                        '  If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then
                        '      wTable.日割日数 = wTable.日割日数 - 1                 ' 07/03/05 V180
                        '    End If                                                  ' 07/03/05 V180
                        'End If                                                      ' 07/03/05 V180
                      End If                                                        ' 07/01/31 V180
                      
                      'If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then
                      '  wTable.日割日数 = wTable.日割日数 - 1                       ' 07/01/31 V180
                      'End If                                                        ' 07/01/31 V180
                   
                      'wTable.利率 = MBD010_金利参照(p借入金マスタ, DateAdd("m", -1, w返済予定年月))
                      
                    Else
                      '利息後払い
                      If p借入金マスタ.利息支払方法 = 0 Then                        ' 07/01/31 V180
                        '利息毎月支払
                        If w返済予定年月 <= p借入金マスタ.金利初回年月 Then         ' 07/01/31 V180
                          wTable.日割日数 = DateDiff("d", p借入金マスタ.実行日, w解約実行日) + 1
                          
                          w内入開始年月日 = p借入金マスタ.実行日          '10/01/30
                          w内入終了年月日 = w解約実行日                   '10/01/30
                           
                          wTable.利息対象期間日数 = wTable.日割日数     '09/12/29
                                                    
                          If p借入金マスタ.利息控除区分 = 1 Or p借入金マスタ.利息控除区分 = 2 Then
                            wTable.日割日数 = wTable.日割日数 - 1                   ' 07/01/31 V180
                          Else                                                      ' 07/01/31 V180
                            If p借入金マスタ.利息控除区分 = 3 Then                  ' 07/01/31 V180
                                wTable.日割日数 = wTable.日割日数 - 2               ' 07/01/31 V180
                            End If                                                  ' 07/01/31 V180
                          End If                                                    ' 07/01/31 V180
                        Else                                                        ' 07/01/31 V180
                          GDate2 = DateAdd("m", -1, w返済予定年月)                  ' 07/01/31 V180
                          GDate2 = MBD010_利息計算年月日(GDate2, p借入金マスタ.支払日, _
                            p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)
                          wTable.日割日数 = DateDiff("d", GDate2, w解約実行日)      ' 07/01/31 V180
                          
                          w内入開始年月日 = GDate2                        '10/01/30
                          w内入終了年月日 = w解約実行日                   '10/01/30
                          
                          
                          GDate2利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                          wTable.利息対象期間日数 = wTable.日割日数     '09/12/29
                          
                          If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then
                            wTable.日割日数 = wTable.日割日数 - 1                   ' 07/01/31 V180
                          End If                                                    ' 07/01/31 V180
                        End If                                                      ' 07/01/31 V180
                      Else                                                          ' 07/01/31 V180
                        '利息一括支払
                        'wTable.日割日数 = DateDiff("d", p借入金マスタ.実行日, w解約実行日) + 1 ' 07/01/31 V180
                        
                        'wTable.利息対象期間日数 = DateDiff("d", p借入金マスタ.実行日, w解約実行日) + 1 'V182 2008/01/29
                        
                        'If p借入金マスタ.利息控除区分 = 1 Or p借入金マスタ.利息控除区分 = 2 Then
                            'wTable.日割日数 = wTable.日割日数 - 1                   ' 07/01/31 V180
                        'Else                                                        ' 07/01/31 V180
                            'If p借入金マスタ.利息控除区分 = 3 Then                  ' 07/01/31 V180
                                'wTable.日割日数 = wTable.日割日数 - 2               ' 07/01/31 V180
                            'End If                                                  ' 07/01/31 V180
                        'End If                                                      ' 07/01/31 V180
                          
                        GDate1 = MBD010_借入金解約据置期間一括利息支払日(w解約実行日, _
                                                  p借入金マスタ.実行日, _
                                                  w返済予定年月, _
                                                  w一括支払利息開始年月, _
                                                  p借入金マスタ.初回返済年月, _
                                                  p借入金マスタ.利息区分, _
                                                  p借入金マスタ.返済単位月数, _
                                                  p借入金マスタ.支払日, _
                                                  p借入金マスタ.営業日区分, _
                                                  p借入金マスタ.利息計算日数区分)    '10/02/02
                          
                          
                        'GDate1 = MBD010_利息計算年月日(GDate1, p借入金マスタ.支払日, _
                        ' p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分) '08/03/14 V185
                         
                         
                        '**一括支払利息開始日までの解約の場合
                        w年月日 = MBD010_利息計算年月日(w一括支払利息開始年月, p借入金マスタ.支払日, _
                         p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分) '09/12/20 V185
                        
                        If Format(w年月日, "yyyy/mm/dd") >= Format(w解約実行日, "yyyy/mm/dd") Then   ' 2010/01/24
                            GDate1利息対象年月日 = DateDiff("d", p借入金マスタ.実行日, w解約実行日) + 1 ' 2009/12/20 V188                  '08/03/14 V185
                            
                            wTable.日割日数 = DateDiff("d", p借入金マスタ.実行日, w解約実行日) + 1  '2009/12/20
                            
                            w内入開始年月日 = p借入金マスタ.実行日          '10/01/30
                            w内入終了年月日 = w解約実行日                   '10/01/30
                            
                            wTable.利息対象期間日数 = wTable.日割日数       '09/12/29
                            w初回利息F = 1                                  '2009/12/20
                         Else                                               ' 2009/12/20
                            wTable.日割日数 = DateDiff("d", GDate1, w解約実行日)     '2009/12/20 V185
                            
                            w内入開始年月日 = GDate1                        '10/01/30
                            w内入終了年月日 = w解約実行日                   '10/01/30
                            
                            wTable.利息対象期間日数 = wTable.日割日数       '09/12/29
                            w初回利息F = 0                                  '2009/12/20
                         End If                                             ' 2009/12/20
                         
                        
                        '***** 仮に無効とした　'*  08/07/21 V188
                        'w初回利息F = 0                  '08/03/16 V185
                        'If Format(GDate1, "yyyymmdd") = Format(p借入金マスタ.実行日, "yyyymmdd") Then '08/03/16 V185
                        '    wTable.日割日数 = DateDiff("d", GDate1, w解約実行日) + 1    '08/03/16 V185
                        '    wTable.利息対象期間日数 = DateDiff("d", GDate1, w解約実行日) + 1 '08/03/16 V185
                        '    w初回利息F = 1                  '08/03/16 V185
                        'Else                                '08/03/16 V185
                        '    wTable.日割日数 = DateDiff("d", GDate1, w解約実行日)        '08/03/14 V185
                        '    GDate1利息対象年月日 = GDate利息対象年月日                  '08/03/14 V185
                        '    wTable.利息対象期間日数 = DateDiff("d", GDate1利息対象年月日, w解約実行日)  '08/03/14 V185
                        'End If                              '08/03/16 V185
                        ''*****
                        
                        '*** 仮に追加した　08/07/21 V188
                        
                        'wTable.日割日数 = DateDiff("d", p借入金マスタ.実行日, w解約実行日) + 1 '08/07/21 V188
                        
                        'GDate1利息対象年月日 = DateDiff("d", p借入金マスタ.実行日, w解約実行日) + 1 ' 08/07/21 V188                  '08/03/14 V185
                        'wTable.利息対象期間日数 = DateDiff("d", p借入金マスタ.実行日, w解約実行日) + 1  '08/03/14 V185
                        'w初回利息F = 1                               ' 08/07/21 V188
                        '***
                        
                        If w初回利息F = 1 And (p借入金マスタ.利息控除区分 = 1 Or p借入金マスタ.利息控除区分 = 3) Then '08/03/16 V185
                            wTable.日割日数 = wTable.日割日数 - 1   '08/03/16 V185
                        End If                                      '08/03/16 V185
                        
                        If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then '08/03/16 V185
                            wTable.日割日数 = wTable.日割日数 - 1   '08/03/16 V185
                        End If                                      '08/03/16 V185
                        
                        
                      End If                                                        ' 07/01/31 V180
                      
                    End If                                                          ' 07/01/31 V180
                         
                    wTable.日割日数 = wTable.日割日数 + w日割日数                   ' 07/03/23 V180
                    
                    'wTable.利率 = MBD010_金利参照(p借入金マスタ, w返済予定年月)
                Else                                                                ' 07/03/23 V180
                    w日割日数 = 0                                                   ' 07/03/23 V180
                End If
            End If
            
            '--------------
            '** 利息計算 **
            '--------------
            '** 解約の時　解約実行日を、実際年月日にセット **
            
            If Not IsNull(w解約実行日) And wTable.実際年月日 >= w解約実行日 Then   '解約
                wTable.実際年月日 = w解約実行日
            End If
            
            '解約日調整＆利息計算年月日　セット（据置期間）
            If Not IsNull(w解約実行日) And _
                Format(wTable.実際年月日, "yyyy/mm/dd") = Format(w解約実行日, "yyyy/mm/dd") Then '10/01/04
                    wTable.利息計算年月日 = w解約実行日         '10/01/23
                
            Else                                            '10/01/04
                wTable.利息計算年月日 = MBD010_利息計算年月日(wTable.返済予定年月, p借入金マスタ.支払日, _
                         p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分) '10/01/04
            End If                                          '10/01/04
            
            '*** 変動金利　設定
            If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then   ' 08/012/10 V189
                If Not IsNull(w解約実行日) And _
                    Format(wTable.実際年月日, "yyyy/mm/dd") = Format(w解約実行日, "yyyy/mm/dd") Then
                    w利息計算基準年月日 = DateAdd("d", -1, wTable.利息計算年月日)   '10/01/23
                Else                                                                '10/01/23
                    w利息計算基準年月日 = wTable.利息計算年月日                     '10/01/23
                End If                                                              '10/01/23
                
            Else                                                                ' 08/12/10 V189
                w利息計算基準年月日 = DateAdd("d", -wTable.利息対象期間日数, wTable.利息計算年月日) '10/01/04
            End If                                                              ' 08/12/10 V189
            wTable.利率 = MBD010_金利参照(p借入金マスタ, w利息計算基準年月日)   ' 08/12/10 V189
            
            
            'If p借入金マスタ.金利計算年間日数 = 0 Then                          ' 07/01/31 V180
            '    wTable.利息額 = Fix(p借入金マスタ.融資金額 * CCur(wTable.利率) * wTable.日割日数 / 36500)
            'Else                                                                    ' 07/01/31 V180
            '    wTable.利息額 = Fix(p借入金マスタ.融資金額 * CCur(wTable.利率) * wTable.日割日数 / 36000)
            'End If                                                                  ' 07/01/31 V180
            wTable.利息額 = MBD010_利息計算小数点5桁(wTable.利率, p借入金マスタ.融資金額, _
                                    wTable.日割日数, p借入金マスタ.金利計算年間日数) '09/12/30
            
            
            wTable.返済金額 = wTable.利息額
            
            '** 解約の時　解約年月日を、実際年月日にセット **
            If Not IsNull(w解約実行日) And w実際年月日 >= w解約実行日 Then   '解約
                wTable.実際年月日 = w解約実行日
            End If
            
            '借入実行日= 据置１回目年月日&利息先払&利息一括支払　データは　実行日のＤＡＴＡを優先する
            If Format(p借入金マスタ.実行日, "yyyy/mm/dd") = Format(w実際年月日, "yyyy/mm/dd") _
                And p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") _
                And p借入金マスタ.利息支払方法 = 1 Then                     ' 08/07/20 V188
                Call MBD010_内入処理(p借入金マスタ)                         ' 08/12/05 V189
            Else                                                            ' 08/07/20 V188
                If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then   ' 08/12/05 V189
                    wTable.据置x回目 = 2                                    ' 09/01/16 V189
                    Call MBD010_借入金テーブルWrite(wTable, p借入金マスタ)  '2016/09/23
                    Call MBD010_内入処理(p借入金マスタ)                     ' 08/12/05 V189
                Else                                                        ' 08/12/05 V189
                    
                    Call MBD010_内入処理(p借入金マスタ)                     ' 08/12/05 V189
                    wTable.据置x回目 = 2                                    ' 09/01/16 V189
                    Call MBD010_借入金テーブルWrite(wTable, p借入金マスタ)  '2016/09/23
                End If                                                      ' 08/12/05 V189
            End If                              ' 08/07/20 V188
            
    Next
  End If                                        '10/05/06 V195
  
    ' -----------------------------------------
    '            初回返済以降のＤＡＴＡ
    ' -----------------------------------------
    w融資残高 = p借入金マスタ.融資金額 'XX Add
        
    For wRecCount = 0 To p借入金マスタ.支払回数 - 1
        
        w返済予定年月 = DateAdd("m", wRecCount, p借入金マスタ.初回返済年月)
        
           
        If wRecCount = 0 Then                               ' 08/07/19 V188
                w実際年月日 = p借入金マスタ.初回返済実行日  ' 08/07/19 V188
        Else                                            ' 08/07/19 V188
            If wRecCount = p借入金マスタ.支払回数 - 1 Then  ' 08/07/19 V188
                w実際年月日 = p借入金マスタ.最終返済実行日  ' 08/07/19 V188
            Else                                        ' 08/07/19 V188
                w実際年月日 = MXA030_翌営業年月日計算(w返済予定年月, p借入金マスタ.支払日, p借入金マスタ.営業日区分) ' 07/01/30 V180
            End If                                      ' 08/07/19 V188
        End If                                          ' 08/07/19 V188
        
        
        'w内入開始年月日 = Null                          ' 08/12/05 V189
        'w内入終了年月日 = Null                          ' 08/12/05 V189
                
        wTable.借入番号 = p借入金マスタ.借入番号
        wTable.返済回数 = wRecCount + 1
        wTable.据置x回目 = 2                    ' 08/12/06 V189
        wTable.返済予定年月 = w返済予定年月
        wTable.実際年月日 = w実際年月日
        OLD実際年月日 = NEW実際年月日                   ' 07/03/24 V180
        NEW実際年月日 = wTable.実際年月日               ' 07/03/24 V180
        
        wTable = MBD010_借入金テーブルRead(wTable)
        
            wTable.実際年月日 = w実際年月日
            
            '**返済単位月数による返済月の算出
            w返済単位判定 = wRecCount + 1 + p借入金マスタ.返済単位月数 - 1
            w返済単位回数 = Fix(w返済単位判定 / p借入金マスタ.返済単位月数)
            If w返済単位回数 * p借入金マスタ.返済単位月数 = w返済単位判定 Then
                Select Case wRecCount
                    Case 0:                     wTable.元金額 = p借入金マスタ.初回返済額
                    Case p借入金マスタ.支払回数 - 1: wTable.元金額 = p借入金マスタ.最終返済額
                    Case Else:                  wTable.元金額 = p借入金マスタ.毎月返済額
                End Select
            Else
                wTable.元金額 = 0
            End If
            
                
                    
            wTable.保証料 = 0
            'wTable.利率 = MBD010_金利参照(p借入金マスタ, w返済予定年月)
            w内入開始年月日 = Null                                          '10/02/07
            w内入終了年月日 = Null                                          '10/02/07
            '***元金、利息がゼロの時　w内入開始年月日　w内入終了年月日 の初期値セット
          'If Format(w返済予定年月, "yyyy/mm/dd") >= Format(p借入金マスタ.最終返済年月, "yyyy/mm/dd") Then
          'Else                                                              '10/02/07
          '
          '  wx = w返済単位判定 - w返済単位回数 * p借入金マスタ.返済単位月数 '10/02/05 後払
          '  wY = p借入金マスタ.返済単位月数                                 '10/02/05
          '  wZ = wY - wx                                                    '10/02/05 先払
          '  If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then       '10/02/05
          '      '*利息先払
          '      GDate1 = DateAdd("m", wZ, w返済予定年月)                                 '10/02/05
          '      GDate1 = MBD010_利息計算年月日(GDate1, p借入金マスタ.支払日, _
          '           p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)      '10/02/05
          '      If Format(DateAdd("m", wZ, w返済予定年月), "yyyy/mm/dd") = _
          '          Format(p借入金マスタ.最終返済年月, "yyyy/mm/dd") _
          '          And Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
          '              Format(p借入金マスタ.最終返済実行日, "yyyy/mm/dd") Then '10/02/07
          '          GDate1 = p借入金マスタ.最終返済実行日                       '10/02/07
          '      End If                                                          '10/02/07
          '
          '      GDate2 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
          '           p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)      '10/02/05
          '  Else                                                                        '10/02/05
          '      '*利息後払
          '      GDate1 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
          '           p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)      '10/02/05
          '      GDate2 = DateAdd("m", -wx, w返済予定年月)                               '10/02/05
          '      GDate2 = MBD010_利息計算年月日(GDate2, p借入金マスタ.支払日, _
          '           p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)      '10/02/05
          '  End If                                                                      '10/02/05
          '  w内入開始年月日 = GDate2                                                    '10/02/05
          '  w内入終了年月日 = GDate1                                                    '10/02/05
          'End If                                                                '10/02/07
          
            '--------------
            '** 日割日数 **
            '--------------
            wTable.日割日数 = 0                                     'V180 07/02/01
            If w返済単位回数 * p借入金マスタ.返済単位月数 <> w返済単位判定 _
                                  And p借入金マスタ.利息支払方法 = 1 Then    '10/01/09                             'V180 07/02/01
            Else                                                    'V180 07/02/01
            
              If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
                '利息先払い
                If p借入金マスタ.利息支払方法 = 0 Then              'V180 07/02/01
                    '利息毎月支払
                    If Format(w返済予定年月, "yyyy/mm/dd") = _
                        Format(p借入金マスタ.最終返済年月, "yyyy/mm/dd") Then   '10/02/10
                        GDate1 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
                            p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)   '10/02/10
                        If Format(GDate1, "yyyy/mm/dd") = Format(w実際年月日, "yyyy/mm/dd") Then '10/02/10
                        Else                            '10/02/10
                            GDate1 = w実際年月日        '10/02/10
                        End If                          '10/02/10
                    
                    
                    Else                                '10/02/10
                        GDate1 = DateAdd("m", 1, w返済予定年月)
                        GDate1 = MBD010_利息計算年月日(GDate1, p借入金マスタ.支払日, _
                        p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)  ' 07/01/30 V180
                     
                     
                        '*最終返済日が手打ちの時の調整
                        If Format(DateAdd("m", 1, w返済予定年月), "yyyy/mm/dd") = Format(p借入金マスタ.最終返済年月, "yyyy/mm/dd") _
                           And Format(GDate利息対象年月日, "yyyy/mm/dd") _
                          <> Format(p借入金マスタ.最終返済実行日, "yyyy/mm/dd") Then  '10/02/10
                            GDate1 = p借入金マスタ.最終返済実行日           '10/02/10
                            GDate利息対象年月日 = p借入金マスタ.最終返済実行日   '10/02/10
                        End If                                              '10/02/10
                    End If                                                  '10/02/10
                     
                     
                     
                     
                     
                    'If wRecCount = sv支払回数 - 2 Then                  ' 08/07/19 V188
                    '    If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                    '       Format(p借入金マスタ.最終返済実行日, "yyyy/mm/dd") Then      ' 08/07/22 V188
                    '        GDate1 = p借入金マスタ.最終返済実行日                       ' 08/07/19 V188
                    '        GDate利息対象年月日 = p借入金マスタ.最終返済実行日         ' 08/07/19 V188
                    '    End If                                                          ' 08/07/22 V188
                    'End If                                                          ' 08/07/19 V188
                     
                    GDate1利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                     
                    GDate2 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
                     p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)  ' 07/01/30 V180
                     
                    If wRecCount = 0 Then                                   ' 08/07/19 V188
                        If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                           Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then      ' 08/07/22 V188
                            GDate2 = p借入金マスタ.初回返済実行日                       ' 08/07/19 V188
                            GDate利息対象年月日 = p借入金マスタ.初回返済実行日         ' 08/07/19 V188
                        End If                                                          ' 08/07/22 V188
                    End If                                                              ' 08/07/19 V188
                     
                    wTable.日割日数 = DateDiff("d", GDate2, GDate1)     'V182 2008/01/29
                    w内入開始年月日 = GDate2                            ' 08/12/05 V189
                    w内入終了年月日 = GDate1                            ' 08/12/05 V189
                    
                    GDate2利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                    wTable.利息対象期間日数 = wTable.日割日数           '09/12/29
                    w利息対象内入開始年月日 = GDate2利息対象年月日          ' 08/12/08 V189
                    w利息対象内入終了年月日 = GDate1利息対象年月日          ' 08/12/08 V189
                    
                    GDate2 = MBD010_利息計算年月日(p借入金マスタ.最終返済年月, p借入金マスタ.支払日, _
                     p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)  ' 07/01/30 V180
                     
                    If p借入金マスタ.実行日 = p借入金マスタ.初回返済実行日 _
                        And w実際年月日 = p借入金マスタ.実行日 _
                        And (p借入金マスタ.利息控除区分 = 1 Or p借入金マスタ.利息控除区分 = 3) Then '10/05/06 V195
                        wTable.日割日数 = wTable.日割日数 - 1       '10/05/06 V195
                    End If                                          '10/05/06 V195
                    
                     
                    GDate1 = DateAdd("m", 1, w返済予定年月)         'V180 07/02/01
                    If GDate1 = p借入金マスタ.最終返済年月 Then     'V180 07/02/01
                        If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then
                            wTable.日割日数 = wTable.日割日数 - 1   'V180 07/02/01
                        End If                                      'V180 07/02/01
                    End If                    'V180 07/02/01
                    
                    
                Else                                                'V180 07/02/01
                    '利息一括支払
                    If w返済予定年月 = p借入金マスタ.最終返済年月 Then  'V180 07/02/01
                    Else                                            'V180 07/02/01
                        GDate1 = DateAdd("m", p借入金マスタ.返済単位月数, w返済予定年月)     'V180 07/02/01
                        GDate1 = MBD010_利息計算年月日(GDate1, p借入金マスタ.支払日, _
                         p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分) ' 07/01/30 V180
                         
                        If wRecCount = sv支払回数 - p借入金マスタ.返済単位月数 - 1 Then ' 08/07/19 V188
                            If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                               Format(p借入金マスタ.最終返済実行日, "yyyy/mm/dd") Then      ' 08/07/22 V188
                                GDate1 = p借入金マスタ.最終返済実行日                  ' 08/07/19 V188
                                GDate利息対象年月日 = p借入金マスタ.最終返済実行日    ' 08/07/19 V188
                            End If                                                     ' 08/07/22 V188
                        End If                                                         ' 08/07/19 V188
                         
                        GDate1利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                        
                        GDate2 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
                         p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分) ' 07/01/30 V180
                         
                        If wRecCount = 0 Then                                               ' 08/07/19 V188
                            If Format(GDate利息対象年月日, "yyyy/mm/dd") <> Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then ' 08/07/19 V188
                                GDate2 = p借入金マスタ.初回返済実行日                       ' 08/07/19 V188
                                GDate利息対象年月日 = p借入金マスタ.初回返済実行日         ' 08/07/19 V188
                            End If                                                          ' 08/07/19 V188
                        End If                                                              ' 08/07/19 V188
                         
                        wTable.日割日数 = DateDiff("d", GDate2, GDate1)                ' 07/01/31 V180
                        w内入開始年月日 = GDate2                            ' 08/12/05 V189
                        w内入終了年月日 = GDate1                            ' 08/12/05 V189
                        
                        GDate2利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                        wTable.利息対象期間日数 = wTable.日割日数           '09/12/29
                        w利息対象内入開始年月日 = GDate2利息対象年月日          ' 08/12/08 V189
                        w利息対象内入終了年月日 = GDate1利息対象年月日          ' 08/12/08 V189
                        
                        If p借入金マスタ.実行日 = p借入金マスタ.初回返済実行日 _
                            And w実際年月日 = p借入金マスタ.実行日 _
                            And (p借入金マスタ.利息控除区分 = 1 Or p借入金マスタ.利息控除区分 = 3) Then '10/05/06 V195
                            wTable.日割日数 = wTable.日割日数 - 1           '10/05/06 V195
                        End If                                              '10/05/06 V195
                        
                        
                        GDate1 = DateAdd("m", p借入金マスタ.返済単位月数, w返済予定年月)  'V180 07/02/16
                        If GDate1 = p借入金マスタ.最終返済年月 Then 'V180 07/02/01
                            If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then
                                wTable.日割日数 = wTable.日割日数 - 1   'V180 07/02/01
                            End If                                  'V180 07/02/01
                        End If                                      'V180 07/02/01
                        
                        
                    End If                                          'V180 07/02/01
                End If                                              'V180 07/02/01
                
                If w返済予定年月 = p借入金マスタ.最終返済年月 Then  'V180 07/02/01
                    wTable.日割日数 = 0                             'V180 07/02/01
                    wTable.利息対象期間日数 = 0                     'V182 2008/01/29
                End If                                              'V180 07/02/01
                
                    
              Else                                                  'V189 07/02/01
                '利息後払い
                If p借入金マスタ.利息支払方法 = 0 Then              'V180 07/02/01
                    '利息毎月支払
                    If w返済予定年月 = p借入金マスタ.金利初回年月 Then  'V180 07/02/01
                        GDate1 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
                         p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分) ' 07/01/30 V180
                         
                        If wRecCount = 0 Then                                         ' 08/07/19 V188
                            If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                               Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then      ' 08/07/22 V188
                                GDate1 = p借入金マスタ.初回返済実行日                 ' 08/07/19 V188
                                GDate利息対象年月日 = p借入金マスタ.初回返済実行日    ' 08/07/19 V188
                            End If                                                    ' 08/07/22 V188
                        End If                                                              ' 08/07/19 V188
                         
                         
                        If p借入金マスタ.実行日 <> p借入金マスタ.初回返済実行日 Then    '10/06/05 V195
                            wTable.日割日数 = DateDiff("d", p借入金マスタ.実行日, GDate1) + 1 'V180 07/02/01
                        Else                                                '10/05/06 V195
                            wTable.日割日数 = DateDiff("d", p借入金マスタ.実行日, GDate1)   '10/05/06 V195
                        End If                                              '10/05/06 V195
                        
                        w内入開始年月日 = p借入金マスタ.実行日              ' 08/12/05 V189
                        w内入終了年月日 = GDate1                            ' 08/12/05 V189
                        
                        GDate1利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                        wTable.利息対象期間日数 = wTable.日割日数           '09/12/29
                        w利息対象内入開始年月日 = p借入金マスタ.実行日          ' 08/12/08 V189
                        w利息対象内入終了年月日 = GDate1利息対象年月日          ' 08/12/08 V189
                        If p借入金マスタ.利息控除区分 = 1 Or p借入金マスタ.利息控除区分 = 3 Then
                            wTable.日割日数 = wTable.日割日数 - 1   'V180 07/02/01
                        End If                                      'V180 07/02/01
                        
                        If (Format(p借入金マスタ.金利初回年月, "yyyy/mm/dd") = _
                            Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd")) _
                            And (Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd") = _
                                Format(p借入金マスタ.最終返済年月, "yyyy/mm/dd")) _
                            And (p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3) Then        '09/12/31
                            wTable.日割日数 = wTable.日割日数 - 1               '09/12/31
                        End If                                                  '09/12/31
                        
                            
                            
                    Else                                            'V180 07/02/01
                        If w返済予定年月 > p借入金マスタ.金利初回年月 Then  'V180 07/02/01
                            GDate1 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
                             p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分) ' 07/01/30 V180
                             
                            If wRecCount = 0 Then                                   ' 08/07/19 V188
                                If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                                   Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then      ' 08/07/22 V188
                                    GDate1 = p借入金マスタ.初回返済実行日               ' 08/07/19 V188
                                    GDate利息対象年月日 = p借入金マスタ.初回返済実行日  ' 08/07/19 V188
                                End If                                                  ' 08/07/22 V188
                            Else                                                    ' 08/07/20 V188
                                If wRecCount = sv支払回数 - 1 Then                  ' 08/07/20 V188
                                    If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                                       Format(p借入金マスタ.最終返済実行日, "yyyy/mm/dd") Then      ' 08/07/22 V188
                                        GDate1 = p借入金マスタ.最終返済実行日           ' 08/07/19 V188
                                        GDate利息対象年月日 = p借入金マスタ.最終返済実行日  ' 08/07/19 V188
                                    End If                                          ' 08/07/22 V188
                                End If                                              ' 08/07/20 V188
                            End If                                                  ' 08/07/19 V188
                             
                            GDate1利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                            
                            GDate2 = DateAdd("m", -1, w返済予定年月)       'V180 07/02/01
                            GDate2 = MBD010_利息計算年月日(GDate2, p借入金マスタ.支払日, _
                             p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)
                             
                            If wRecCount = 1 Then                         ' 08/07/19 V188
                                If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                                   Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then      ' 08/07/22 V188
                                    GDate2 = p借入金マスタ.初回返済実行日                       ' 08/07/19 V188
                                    GDate利息対象年月日 = p借入金マスタ.初回返済実行日          ' 08/07/19 V188
                                End If                                                          ' 08/07/22 v188
                            End If                                                          ' 08/07/19 V188
                             '
                            wTable.日割日数 = DateDiff("d", GDate2, GDate1) 'V180 07/02/01
                            w内入開始年月日 = GDate2                            ' 08/12/05 V189
                            w内入終了年月日 = GDate1                            ' 08/12/05 V189
                            
                            GDate2利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                            wTable.利息対象期間日数 = wTable.日割日数           '09/12/29
                            w利息対象内入開始年月日 = GDate2利息対象年月日          ' 08/12/08 V189
                            w利息対象内入終了年月日 = GDate1利息対象年月日          ' 08/12/08 V189
                            If w返済予定年月 = p借入金マスタ.最終返済年月 Then      'V180 07/02/01
                                If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then      'V180 07/02/01
                                    wTable.日割日数 = wTable.日割日数 - 1           'V180 07/02/01
                                End If                                              'V180 07/02/01
                            End If                                                  'V180 07/02/01
                                    
                                
                        End If                                      'V180 07/02/01
                    End If                                          'V180 07/02/01
                    
                Else                                                'V180 07/02/01
                    '利息一括支払
                    If p借入金マスタ.初回返済年月 = p借入金マスタ.最終返済年月 _
                        And Format(w一括支払利息開始年月, "yyyy/mm/dd") = _
                            Format(p借入金マスタ.最終返済年月, "yyyy/mm/dd") Then 'V180 07/02/01
                        If w返済予定年月 = p借入金マスタ.最終返済年月 Then  'V180 07/02/01
                            GDate1 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
                             p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)
                             
                            GDate1 = p借入金マスタ.初回返済実行日       ' 08/07/20 V188
                            GDate利息対象年月日 = p借入金マスタ.初回返済実行日  '08/07/20 V188
                            
                            If p借入金マスタ.実行日 <> p借入金マスタ.初回返済実行日 Then    '10/05/06 V195
                                wTable.日割日数 = DateDiff("d", p借入金マスタ.実行日, GDate1) + 1 'V180
                            Else                                                '10/05/06 V195
                                wTable.日割日数 = DateDiff("d", p借入金マスタ.実行日, GDate1)   '10/05/06 V195
                            End If                                              '10/05/06 V195
                            
                            w内入開始年月日 = p借入金マスタ.実行日              ' 08/12/05 V189
                            w内入終了年月日 = GDate1                            ' 08/12/05 V189
                            
                            
                            GDate1利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                            wTable.利息対象期間日数 = wTable.日割日数           '09/12/29
                            w利息対象内入開始年月日 = p借入金マスタ.実行日          ' 08/12/08 V189
                            w利息対象内入終了年月日 = GDate1利息対象年月日          ' 08/12/08 V189
                            
                            If p借入金マスタ.利息控除区分 = 1 Or p借入金マスタ.利息控除区分 = 2 Then
                                wTable.日割日数 = wTable.日割日数 - 1   'V180 07/02/01
                            Else                                        'V180 07/02/01
                                If p借入金マスタ.利息控除区分 = 3 Then  'V180 07/02/01
                                    wTable.日割日数 = wTable.日割日数 - 2   'V180 07/02/01
                                End If                                  'V180 07/02/01
                            End If                                      'V180 07/02/01
                        End If                                          'V180 07/02/01
                        
                    Else                                                'V180 07/02/01
                        'If w返済予定年月 = p借入金マスタ.初回返済年月 Then      'V180 07/02/01
                        '    GDate1 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
                        '      p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分) 'V180
                        '    wTable.日割日数 = DateDiff("d", p借入金マスタ.実行日, GDate1) + 1 'V180
                        '
                        '    GDate1利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                        '    wTable.利息対象期間日数 = DateDiff("d", p借入金マスタ.実行日, GDate1利息対象年月日) + 1
                            
                        '    If p借入金マスタ.利息控除区分 = 1 Or p借入金マスタ.利息控除区分 = 3 Then    ' 07/02/17 V180
                        '        wTable.日割日数 = wTable.日割日数 - 1
                        '    End If                                      'V180 07/02/01
                        'Else                                            'V180 07/02/01
                        
                        
                            GDate1 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
                                p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分) 'V180
                                
                            If wRecCount = 0 Then                                   ' 08/07/19 V188
                                If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                                   Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then      ' 08/07/22 V188
                                    GDate1 = p借入金マスタ.初回返済実行日               ' 08/07/19 V188
                                    GDate利息対象年月日 = p借入金マスタ.初回返済実行日  ' 08/07/19 V188
                                End If                                                  ' 08/07/22 V188
                            Else                                                    ' 08/07/20 V188
                                If wRecCount = (sv支払回数 - 1) Then                ' 08/07/20 V188
                                    If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                                       Format(p借入金マスタ.最終返済実行日, "yyyy/mm/dd") Then      ' 08/07/22 V188
                                        GDate1 = p借入金マスタ.最終返済実行日           ' 08/07/20 V188
                                        GDate利息対象年月日 = p借入金マスタ.最終返済実行日  ' 08/07/20 V188
                                    End If                                              ' 08/07/22 V188
                                 End If                                             ' 08/07/20 V188
                            End If                                                  ' 08/07/20 V188
                                 
                            GDate1利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                            
                            
                                
                            GDate2 = DateAdd("m", -p借入金マスタ.返済単位月数, w返済予定年月) 'V180
                            GDate2 = MBD010_利息計算年月日(GDate2, p借入金マスタ.支払日, _
                              p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)
                              
                            If wRecCount = 0 And (Format(GDate2, "yyyy/mm/dd") < Format(p借入金マスタ.実行日, "yyyy/mm/dd") _
                              Or Format(w一括支払利息開始年月, "yyyy/mm/dd") = Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd")) _
                               Then ' 2009/12/20                               ' 08/07/19 V188
                                
                                If p借入金マスタ.実行日 <> p借入金マスタ.初回返済実行日 Then    '10/05/06 V195
                                    GDate2 = DateAdd("d", -1, p借入金マスタ.実行日)     ' 08/07/19 V188
                                    GDate利息対象年月日 = DateAdd("d", -1, p借入金マスタ.実行日) ' 08/07/19 V188
                                Else                                                    '10/05/06 V195
                                    GDate2 = p借入金マスタ.実行日       '10/05/06 V195
                                    GDate利息対象年月日 = p借入金マスタ.実行日      '10/05/06 V195
                                End If                                              '10/05/06 V195
                                
                            Else                                                    ' 08/07/20 V188
                                If 0 = (wRecCount - p借入金マスタ.返済単位月数) Then ' 08/07/20 V188
                                    If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                                       Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then      ' 08/07/22 V188
                                        GDate2 = p借入金マスタ.初回返済実行日           ' 08/07/20 V188
                                        GDate利息対象年月日 = p借入金マスタ.初回返済実行日  ' 08/07/20 V188
                                    End If                                          ' 08/07/22 V188
                                End If                                              ' 08/07/20 V188
                            End If
                              
                            wTable.日割日数 = DateDiff("d", GDate2, GDate1) ' 08/07/20 V188
                            w内入開始年月日 = GDate2                            ' 08/12/05 V189
                            w内入終了年月日 = GDate1                            ' 08/12/05 V189
                            
                            GDate2利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                            wTable.利息対象期間日数 = wTable.日割日数           '09/12/29
                            w利息対象内入開始年月日 = GDate2利息対象年月日          ' 08/12/08 V189
                            w利息対象内入終了年月日 = GDate1利息対象年月日          ' 08/12/08 V189
                            If w返済予定年月 = p借入金マスタ.最終返済年月 Then  'V180 07/02/01
                                If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then
                                    wTable.日割日数 = wTable.日割日数 - 1   'V180 07/02/01
                                End If                                      'V180 07/02/01
                            End If                                          'V180 07/02/01
                            
                            '***初回返済日　が　金利初回返済日　の時　実行日の控除処理
                            If (wRecCount = 0 And _
                               (Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd") = Format(p借入金マスタ.金利初回年月, "yyyy/mm/dd")) _
                               And (p借入金マスタ.利息控除区分 = 1 Or p借入金マスタ.利息控除区分 = 3)) _
                               Or _
                               (p借入金マスタ.実行日 = p借入金マスタ.初回返済実行日 _
                                And w返済予定年月 = p借入金マスタ.金利初回年月 _
                                And (p借入金マスタ.利息控除区分 = 1 Or p借入金マスタ.利息控除区分 = 3)) Then   '10/05/06 V195
                               wTable.日割日数 = wTable.日割日数 - 1    '09/12/26
                            End If                                      '09/12/26
                            
                        'End If                                              'V180 07/02/01
                        
                        
                    End If                                              'V180 07/02/01
                End If                                                  'V180 07/02/01
                
              End If                                                    'V180 07/02/01
              
            End If                                                      'V180 07/02/01
            
                    
            '** 解約 **
            If Not IsNull(w解約実行日) Then
                If w実際年月日 >= w解約実行日 Then
                    w差月数 = DateDiff("M", OLD実際年月日, NEW実際年月日)           ' 07/03/24 V180
                    If w実際年月日 = w解約実行日 And w差月数 >= 2 And p借入金マスタ.営業日区分 = 0 Then     ' 07/03/24 V180
                        w日割日数 = wTable.日割日数                                 ' 07/03/23 V180
                    End If                                                          ' 07/03/23 V180
                    
                    wTable.元金額 = 0
                    If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
                        '利息先払い
                        If p借入金マスタ.利息支払方法 = 0 Then          'V180 07/02/01
                            '利息毎月支払
                            GDate1 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
                              p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)
                              
                            If wRecCount = 0 Then                       ' 08/07/20 V188
                                If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                                   Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then      ' 08/07/22 V188
                                    GDate1 = p借入金マスタ.初回返済実行日   ' 08/07/20 V188
                                    GDate利息対象年月日 = p借入金マスタ.初回返済実行日  ' 08/07/20 V188
                                End If                                  ' 08/07/22 V188
                            End If                                      ' 08/07/20 V188
                            
                            
                            '*最終返済日が手打ちの時の調整
                            If Format(w返済予定年月, "yyyy/mm/dd") = Format(p借入金マスタ.最終返済年月, "yyyy/mm/dd") _
                                And Format(GDate利息対象年月日, "yyyy/mm/dd") _
                                <> Format(p借入金マスタ.最終返済実行日, "yyyy/mm/dd") Then  '10/02/10
                                GDate1 = p借入金マスタ.最終返済実行日           '10/02/10
                                GDate利息対象年月日 = p借入金マスタ.最終返済実行日   '10/02/10
                            End If                                              '10/02/10
                            
                            
                            
                            
                            
                            
                            'If wRecCount = sv支払回数 - 1 Then          ' 08/07/20 V188
                            '    If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                            '       Format(p借入金マスタ.最終返済実行日, "yyyy/mm/dd") Then      ' 08/07/22 V188
                            '        GDate1 = p借入金マスタ.最終返済実行日   ' 08/07/20 V188
                            '        GDate利息対象年月日 = p借入金マスタ.最終返済実行日  ' 08/07/20 V188
                            '    End If                                  ' 08/07/22 V188
                            'End If                                      ' 08/07/20 V188
                              
                            wTable.日割日数 = DateDiff("d", w解約実行日, GDate1) * -1  'V180 07/02/01
                            
                            w内入開始年月日 = w解約実行日                   '10/01/30
                            w内入終了年月日 = GDate1                        '10/01/30
                            
                            GDate1利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                            wTable.利息対象期間日数 = wTable.日割日数           '09/12/29
                            
                            If w返済予定年月 <> p借入金マスタ.最終返済年月 Then     ' 07/03/06 V180
                                If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then 'V180
                                    wTable.日割日数 = wTable.日割日数 - 1       '2012/03/21
                                End If
                            Else                                                '10/02/07
                                'If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then    '10/02/07
                                '    wTable.日割日数 = wTable.日割日数 + 1       '10/02/07
                                'End If                                          'V180 07/03/05
                            End If                                              ' 07/03/06 V180
                            
                        Else                                            'V180 07/02/01
                            '利息一括支払
                            If p借入金マスタ.初回返済年月 = p借入金マスタ.最終返済年月 Then 'V180
                                GDate1 = MBD010_利息計算年月日(w返済予定年月, p借入金マスタ.支払日, _
                                 p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)
                                 
                                If wRecCount = 0 Then                       ' 08/07/21 V188
                                    If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                                       Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then      ' 08/07/22 V188
                                        GDate1 = p借入金マスタ.初回返済実行日
                                        GDate利息対象年月日 = p借入金マスタ.初回返済実行日 ' 08/07/21 V188
                                    End If                                  ' 08/07/22 V188
                                End If             ' 08/07/21 V188
                                                              
                                If wRecCount >= (sv支払回数 - p借入金マスタ.返済単位月数 + 1 - 1) _
                                   And wRecCount <= (sv支払回数 - 1) Then       ' 08/07/21 V188
                                   If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                                      Format(p借入金マスタ.最終返済実行日, "yyyy/mm/dd") Then      ' 08/07/22 V188
                                        GDate1 = p借入金マスタ.最終返済実行日       ' 08/07/21 V188
                                        GDate利息対象年月日 = p借入金マスタ.最終返済実行日  ' 08/07/21 V188
                                   End If                                       ' 08/07/22 V188
                                End If                                          ' 08/07/21 V188
                                
                                    
                                wTable.日割日数 = DateDiff("d", w解約実行日, GDate1) * -1
                                
                                GDate2 = MBD010_利息計算年月日(DateAdd("m", -p借入金マスタ.返済単位月数, p借入金マスタ.最終返済年月), _
                                     p借入金マスタ.支払日, p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)    '10/02/11
                                
                                
                                If Format(GDate2, "yyyy/mm/dd") _
                                          >= Format(w解約実行日, "yyyy/mm/dd") _
                                     Or wTable.日割日数 > 0 Then           ' 10/02/13
                                    If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then '10/02/13
                                        wTable.日割日数 = wTable.日割日数 - 1      '2012/03/21
                                    End If                                          '10/02/13
                                Else                                                '10/02/07
                                    'If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then    '10/02/07
                                    '    wTable.日割日数 = wTable.日割日数 + 1       '10/02/07
                                    'End If                                          '10/02/07
                                End If                                              ' 10/02/13
                                 
                                
                                'If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then    '10/02/13
                                '    wTable.日割日数 = wTable.日割日数 + 1               '10/02/13
                                'End If                                                  '10/02/13
                                
                                
                                w内入開始年月日 = w解約実行日                   '10/01/30
                                w内入終了年月日 = GDate1                        '10/01/30
                                
                                GDate1利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                                wTable.利息対象期間日数 = wTable.日割日数           '09/12/29
                                
                            Else                                        'V180 07/02/01
                                w調整月数 = (wTable.返済回数 - 1) Mod p借入金マスタ.返済単位月数 'V180
                                If w調整月数 = 0 Then                               'V180 07/02/01
                                Else                                                'V180 07/02/01
                                    w調整月数 = p借入金マスタ.返済単位月数 - w調整月数  'V180 07/02/01
                                End If                                              'V180 07/02/01
                                    
                                GDate1 = DateAdd("m", w調整月数, w返済予定年月)     'V180 07/02/01
                                GDate1 = MBD010_利息計算年月日(GDate1, p借入金マスタ.支払日, _
                                   p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)
                                     
                                If wRecCount = 0 Then                       ' 08/07/21 V188
                                    If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                                           Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then  ' 08/07/22 V188
                                            GDate1 = p借入金マスタ.初回返済実行日
                                            GDate利息対象年月日 = p借入金マスタ.初回返済実行日 ' 08/07/21 V188
                                    End If                                  ' 08/07/22 V188
                                End If             ' 08/07/21 V188
                                                              
                                If wRecCount >= (sv支払回数 - p借入金マスタ.返済単位月数 + 1 - 1) _
                                        And wRecCount <= (sv支払回数 - 1) Then       ' 08/07/21 V188
                                    If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                                        Format(p借入金マスタ.最終返済実行日, "yyyy/mm/dd") Then      ' 08/07/22 V188
                                        GDate1 = p借入金マスタ.最終返済実行日       ' 08/07/21 V188
                                        GDate利息対象年月日 = p借入金マスタ.最終返済実行日  ' 08/07/21 V188
                                    End If                                      ' 08/07/2 V188
                                End If                                          ' 08/07/21 V188
                                    
                                wTable.日割日数 = DateDiff("d", w解約実行日, GDate1) * -1 'V180
                                    
                                w内入開始年月日 = w解約実行日                   '10/01/30
                                w内入終了年月日 = GDate1                        '10/01/30
                                    
                                GDate1利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                                wTable.利息対象期間日数 = wTable.日割日数           '09/12/29
                                    
                                    
                                GDate2 = MBD010_利息計算年月日(DateAdd("m", -p借入金マスタ.返済単位月数, p借入金マスタ.最終返済年月), _
                                     p借入金マスタ.支払日, p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)    '10/02/11
                                    
                                    
                                If Format(GDate2, "yyyy/mm/dd") _
                                          >= Format(w解約実行日, "yyyy/mm/dd") _
                                     Or wTable.日割日数 > 0 Then           ' 10/02/11
                                    If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then 'V180
                                        wTable.日割日数 = wTable.日割日数 - 1       '2012/03/21
                                    End If                                          'V180 07/03/06
                                Else                                                '10/02/07
                                    'If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then    '10/02/07
                                    '    wTable.日割日数 = wTable.日割日数 + 1       '10/02/07
                                    'End If                                          '10/02/07
                                End If                                              ' 07/03/06 V180
                                
                            
                            End If                                      'V180 07/02/01
                        End If                                          'V180 07/02/01
                        
                        'If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then 'V180
                        '   wTable.日割日数 = wTable.日割日数 - 1       'V180 07/02/01
                        'End If                                          'V180 07/02/01
                        
                    Else                                                'V180 07/02/01
                        '利息後払い
                        If p借入金マスタ.利息支払方法 = 0 Then          'V180 07/02/01
                            '利息毎月支払
                            If w返済予定年月 <= p借入金マスタ.金利初回年月 Then  'V180 07/02/01
                                If p借入金マスタ.実行日 <> p借入金マスタ.初回返済実行日 Then    '10/05/06 V195
                                    wTable.日割日数 = DateDiff("d", p借入金マスタ.実行日, w解約実行日) + 1
                                Else                                    '10/05/06 V195
                                    wTable.日割日数 = DateDiff("d", p借入金マスタ.実行日, w解約実行日)  '10/05/06 V195
                                End If                                  '10/05/06 V195
                                
                                w内入開始年月日 = p借入金マスタ.実行日          '10/01/30
                                w内入終了年月日 = w解約実行日                   '10/01/30
                                
                                wTable.利息対象期間日数 = wTable.日割日数           '09/12/29
                                
                                If p借入金マスタ.利息控除区分 = 1 Or p借入金マスタ.利息控除区分 = 2 Then
                                    wTable.日割日数 = wTable.日割日数 - 1
                                Else                                    'V180 07/02/01
                                    If p借入金マスタ.利息控除区分 = 3 Then      'V180 07/02/01
                                        wTable.日割日数 = wTable.日割日数 - 2   'V180 07/02/01
                                    End If                                      'V180 07/02/01
                                End If                                          'V180 07/02/01
                            Else                                        'V180 07/02/01
                                GDate2 = DateAdd("m", -1, w返済予定年月)        'V180 07/02/01
                                GDate2 = MBD010_利息計算年月日(GDate2, p借入金マスタ.支払日, _
                                     p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)
                                     
                                If wRecCount = 1 Then                         ' 08/07/19 V188
                                    If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                                       Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then      ' 08/07/22 V188
                                        GDate2 = p借入金マスタ.初回返済実行日               ' 08/07/19 V188
                                        GDate利息対象年月日 = p借入金マスタ.初回返済実行日  ' 08/07/19 V188
                                    End If                                              ' 08/07/22 V188
                                End If                                                  ' 08/07/19 V188
                                     
                                wTable.日割日数 = DateDiff("d", GDate2, w解約実行日)    'V180 07/02/01
                                
                                w内入開始年月日 = GDate2                        '10/01/30
                                w内入終了年月日 = w解約実行日                   '10/01/30
                                
                                GDate2利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                                wTable.利息対象期間日数 = wTable.日割日数           '09/12/29
                                
                                If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then
                                    wTable.日割日数 = wTable.日割日数 - 1       'V180 07/02/01
                                End If                                          'V180 07/02/01
                            End If                                              'V180 07/02/01
                        Else                                                    'V180 07/02/01
                            '利息一括支払
                            w調整月数 = (wTable.返済回数 - 1) Mod p借入金マスタ.返済単位月数 'V180
                            If w調整月数 = 0 Then                               'V180 07/02/01
                                w調整月数 = p借入金マスタ.返済単位月数          'V180 07/02/01
                            End If
                            
                            GDate1 = DateAdd("m", -w調整月数, w返済予定年月)     'V180 07/02/01
                            
                            GDate1 = MBD010_利息計算年月日(GDate1, p借入金マスタ.支払日, _
                                     p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)
                                     
                            '***初回返済以降の最終期間で解約　20009/12/20
                            w年月日 = MBD010_利息計算年月日(w一括支払利息開始年月, p借入金マスタ.支払日 _
                                        , p借入金マスタ.営業日区分, p借入金マスタ.金利計算年間日数) '09/12/10
                                        
                            If wRecCount = 0 And Format(w解約実行日, "yyyy/mm/dd") _
                                            <= Format(w年月日, "yyyy/mm/dd") Then    '10/01/24
                                If p借入金マスタ.実行日 <> p借入金マスタ.初回返済実行日 Then    '10/05/06 V195
                                    GDate1 = DateAdd("d", -1, p借入金マスタ.実行日) ' 08/07/21 V188
                                    GDate利息対象年月日 = DateAdd("d", -1, p借入金マスタ.実行日) ' 08/07/21 V188
                                Else                                '10/05/06 V195
                                    GDate1 = p借入金マスタ.実行日   '10/05/06 V195
                                    GDate利息対象年月日 = p借入金マスタ.実行日  '10/05/06 V195
                                End If                              '10/05/06 V195
                                
                            End If                                  ' 08/07/21 V188
                            
                            If wRecCount >= (2 - 1) _
                                And wRecCount <= (2 + p借入金マスタ.返済単位月数 - 1 - 1) Then ' 08/07/21 V188
                                If Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                                   Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then      ' 08/07/22 V188
                                    GDate1 = p借入金マスタ.初回返済実行日       ' 08/07/21 V188
                                    GDate利息対象年月日 = p借入金マスタ.初回返済実行日  ' 08/07/21 V188
                                End If                                      ' 08/07/22 V188
                            End If                                          ' 08/07/21 V188
                            
                                
                            
                                     
                            wTable.日割日数 = DateDiff("d", GDate1, w解約実行日) 'V180 07/02/01
                            
                            w内入開始年月日 = GDate1                        '10/01/30
                            w内入終了年月日 = w解約実行日                   '10/01/30
                            
                            GDate1利息対象年月日 = GDate利息対象年月日          'V182 2008/01/29
                            wTable.利息対象期間日数 = wTable.日割日数           '09/12/29
                            
                            
                            If ((p借入金マスタ.実行日 = p借入金マスタ.初回返済実行日 _
                                And GDate1 = p借入金マスタ.実行日) _
                                Or _
                                (p借入金マスタ.実行日 <> p借入金マスタ.初回返済実行日 _
                                 And DateAdd("D", 1, GDate1) = p借入金マスタ.実行日)) _
                                And _
                                (p借入金マスタ.利息控除区分 = 1 Or p借入金マスタ.利息控除区分 = 3) Then '10/05/06 V195
                                wTable.日割日数 = wTable.日割日数 - 1       '10/05/06 V195
                            End If                                          '10/05/06 V195
                            
                            If p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3 Then
                                wTable.日割日数 = wTable.日割日数 - 1           'V180 07/02/01
                            End If                                              'V180 07/02/01
                        End If                                                  'V180 07/02/01
                        
                    End If                                                      'V180 07/02/01
                    
                    'wTable.日割日数 = wTable.日割日数 + w日割日数               ' 08/07/25 V188DEL
                    'wTable.利息対象期間日数 = wTable.利息対象期間日数 + w日割日数 08/07/25 V188DEL
                    
                Else                                                            ' 07/03/25 V180
                    w日割日数 = 0                                               ' 07/03/23 V180
                    
                End If
            End If
           
            '--------------
            '** 利息計算 **
            '--------------
            If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息後払") Then   ' 08/12/06 V189
                Call MBD010_内入処理(p借入金マスタ)                                 ' 08/12/06 V189
            End If                                                                  ' 08/12/06 V189
            
            '*** プロジェクト管理で　過去,未来　共に　元金額を＝０と見なす（内入のみ計算対象とする)
            'If p借入金マスタ.プロジェクト番号 > "" Then     ' 09/03/01 V189
            '    'Format(w本日年月日, "yyyy/mm/dd") > Format(wTable.実際年月日, "yyyy/mm/dd") Then ' 08/12/06 V189
            '    wTable.元金額 = 0                               ' 08/12/06 V189
            'End If                                              ' 08/12/06 V189
            
            
            '***
            w融資残高 = w融資残高 - wTable.元金額
            wTable.融資残高 = w融資残高
            
            
            
            '** 解約の時　解約実行日を、実際年月日にセット **
            If Not IsNull(w解約実行日) And wTable.実際年月日 >= w解約実行日 Then   '解約
                wTable.実際年月日 = w解約実行日
            End If
              
            '***解約日調整＆利息計算年月日　のセット　（初回返済年月以降）
            If Not IsNull(w解約実行日) And _
                Format(wTable.実際年月日, "yyyy/mm/dd") = Format(w解約実行日, "yyyy/mm/dd") Then '10/01/04
                wTable.利息計算年月日 = w解約実行日                 '10/01/04
            Else                                                    '10/01/04
                wTable.利息計算年月日 = MBD010_利息計算年月日(wTable.返済予定年月, p借入金マスタ.支払日, _
                         p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分) '10/01/04
                '***初回返済年月日　手打ち
                If Format(wTable.返済予定年月, "yyyy/mm/dd") = Format(p借入金マスタ.初回返済年月, "yyyy/mm/dd") _
                    And Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                        Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then  '10/01/04
                    wTable.利息計算年月日 = p借入金マスタ.初回返済実行日        '10/01/04
                Else                                                            '10/01/04
                    If Format(wTable.返済予定年月, "yyyy/mm/dd") = Format(p借入金マスタ.最終返済年月, "yyyy/mm/dd") _
                        And Format(GDate利息対象年月日, "yyyy/mm/dd") <> _
                            Format(p借入金マスタ.最終返済実行日, "yyyy/mm/dd") Then  '10/01/04
                        wTable.利息計算年月日 = p借入金マスタ.最終返済実行日        '10/01/04
                    End If                                                      '10/01/04
                End If                                                          '10/01/04
            End If                                                              '10/01/04
            
            '*** 変動金利　設定
            If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then   ' 08/012/10 V189
                If Not IsNull(w解約実行日) And _
                    Format(wTable.実際年月日, "yyyy/mm/dd") = Format(w解約実行日, "yyyy/mm/dd") Then
                    w利息計算基準年月日 = DateAdd("d", -1, wTable.利息計算年月日)   '10/01/23
                Else                                                                '10/01/23
                    w利息計算基準年月日 = wTable.利息計算年月日                     '10/01/23
                End If                                                              '10/01/23
                
            Else                                                                ' 08/12/10 V189
                w利息計算基準年月日 = DateAdd("d", -wTable.利息対象期間日数, wTable.利息計算年月日) '10/01/04
            End If                                                              ' 08/12/10 V189
            wTable.利率 = MBD010_金利参照(p借入金マスタ, w利息計算基準年月日)   ' 08/12/10 V189
            
            If p借入金マスタ.実行日 = p借入金マスタ.初回返済実行日 _
                And p借入金マスタ.実行日 = w利息計算基準年月日 _
                And wTable.利息対象期間日数 = 0 Then                            '10/05/06 V195
                wTable.利率 = p借入金マスタ.利率                                '10/05/06 V195
            End If                                                              '10/05/06 V195
            
            
            If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
                '利息先払い
                'If p借入金マスタ.金利計算年間日数 = 0 Then                  'V180 07/02/01
                '    wTable.利息額 = Fix(wTable.融資残高 * CCur(wTable.利率) * wTable.日割日数 / 36500)
                'Else                                                            'V180 07/02/01
                '    wTable.利息額 = Fix(wTable.融資残高 * CCur(wTable.利率) * wTable.日割日数 / 36000)
                'End If                                                          'V180 07/02/01
                wTable.利息額 = MBD010_利息計算小数点5桁(wTable.利率, p借入金マスタ.融資残高, _
                                    wTable.日割日数, p借入金マスタ.金利計算年間日数) '10/06/15
            
                
            Else
                '利息後払い
                'If p借入金マスタ.金利計算年間日数 = 0 Then                  'V180 07/02/01
                '    wTable.利息額 = Fix((wTable.融資残高 + wTable.元金額) * CCur(wTable.利率) * wTable.日割日数 / 36500)   '利息後払い
                'Else
                '    wTable.利息額 = Fix((wTable.融資残高 + wTable.元金額) * CCur(wTable.利率) * wTable.日割日数 / 36000)   '利息後払い
                'End If
                wTable.利息額 = MBD010_利息計算小数点5桁(wTable.利率, wTable.融資残高 + wTable.元金額, _
                                    wTable.日割日数, p借入金マスタ.金利計算年間日数) '10/06/15
                
            End If
            wTable.返済金額 = wTable.利息額 + wTable.元金額
                     
            '** 解約の時　解約実行日を、実際年月日にセット **
            'If Not IsNull(w解約実行日) And wTable.実際年月日 >= w解約実行日 Then   '解約
            '    wTable.実際年月日 = w解約実行日
            'End If
            
        wTable.据置x回目 = 2                                    ' 09/01/16 V189
        Call MBD010_借入金テーブルWrite(wTable, p借入金マスタ)  '2016/09/23
        
        If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then   ' 08/12/06 V189
            Call MBD010_内入処理(p借入金マスタ)                                 ' 08/12/06 V189
        End If                                                                  ' 08/12/06 V189
        
    Next
            
    ' -----------------------------------------
    '                保証料算出
    ' -----------------------------------------
    w解約保証料戻 = 0
    
    If p借入金マスタ.保証料率 <> 0 Then
    
        w保証料支払日 = 18

        wRet保証料算出 = MBD010_保証料算出(p借入金マスタ.借入番号, _
                                          p借入金マスタ.自己資金フラグ, _
                                          p借入金マスタ.有担保フラグ, _
                                          p借入金マスタ.保証料分割フラグ, _
                                          p借入金マスタ.保証料率, _
                                          p借入金マスタ.融資金額, _
                                          p借入金マスタ.実行日, _
                                          w解約実行日, _
                                          p借入金マスタ.初回返済実行日, _
                                          p借入金マスタ.最終返済実行日, _
                                          p借入金マスタ.初回返済年月, _
                                          p借入金マスタ.最終返済年月, _
                                          p借入金マスタ.支払回数, _
                                          p借入金マスタ.据置回数, _
                                          w実行支払年月, _
                                          p借入金マスタ.営業日区分 _
                                          )
        
        w解約保証料戻 = wRet保証料算出.解約保証料戻
         
         
        '保証料通常登録
        wTable.借入番号 = p借入金マスタ.借入番号
        wTable.返済回数 = 0
        wTable.据置x回目 = 5                    ' 08/12/16 V189
        'wTable.返済予定年月 = p借入金マスタ.実行日
        wTable.実際年月日 = p借入金マスタ.実行日
        wTable.元金額 = 0                       ' 08/12/14 V189
        
        wTable = MBD010_借入金テーブルRead(wTable)
        '    wTable.金融保証料 = wRet保証料算出.初回保証料
            wTable.保証料 = wRet保証料算出.初回保証料
        Call MBD010_借入金テーブルWrite(wTable, p借入金マスタ)      '2016/09/23
   
        '保証料通常登録(分割時)
                    
        For wCount = 1 To 9  'X年後
            If wRet保証料算出.保証料X年後(wCount) <> 0 Then
                    
                w返済予定年月 = DateSerial((Year(p借入金マスタ.実行日) + wCount), Month(p借入金マスタ.実行日), 1)
                w実際予定年月日 = MXA030_翌営業年月日計算(w返済予定年月, w保証料支払日, p借入金マスタ.営業日区分) ' 07/01/30 V180
                GDate1 = MXA030_翌営業年月日計算(w返済予定年月, w保証料支払日, p借入金マスタ.営業日区分) ' 07/01/30 V180
            '    w実際予定年月日 = DateSerial((Year(p借入金マスタ.実行日) + wCount), Month(p借入金マスタ.実行日), w保証料支払日)　2003/7/3
            '    GDate1 = DateSerial(Year(w返済予定年月), Month(w返済予定年月), w保証料支払日) '2003/7/3
             
                If w解約実行日 > GDate1 _
                Or IsNull(w解約実行日) Then
        
                    wTable.借入番号 = p借入金マスタ.借入番号
                    wTable.返済回数 = 0
                    wTable.据置x回目 = 5
                    wTable.返済予定年月 = w返済予定年月
                      
                    wTable = MBD010_借入金テーブルRead(wTable)
                        wTable.保証料 = wRet保証料算出.保証料X年後(wCount)
                    '    wTable.金融保証料 = wRet保証料算出.保証料X年後(wCount)
                        wTable.実際年月日 = w実際予定年月日
                        wTable.返済予定年月 = w返済予定年月
                        wTable.元金額 = 0
                        wTable.日割日数 = 0
                        wTable.利息対象期間日数 = 0 'V182 2008/01/29
                        wTable.返済金額 = 0
                        wTable.融資残高 = 0
                        wTable.利息額 = 0
                        wTable.利率 = 0
                         
                    Call MBD010_借入金テーブルWrite(wTable, p借入金マスタ)      '2016/09/23
                End If
            End If
        Next
   
        '解約保証料戻
        If w解約保証料戻 <> 0 Then
            wTable.借入番号 = p借入金マスタ.借入番号
            wTable.返済回数 = 0
            wTable.据置x回目 = 5
            wTable.返済予定年月 = w解約実行日
            wTable.実際年月日 = w解約実行日
            
            wTable = MBD010_借入金テーブルRead(wTable)
                wTable.保証料 = w解約保証料戻
            '    wTable.金融保証料 = -1 * w解約保証料戻
                wTable.実際年月日 = w解約実行日
               
            Call MBD010_借入金テーブルWrite(wTable, p借入金マスタ)          '2016/09/23
        End If
        
    End If
      
    MBD010_借入金テーブル作成 = w解約保証料戻
    
    Call MBD010_融資残高再計算(p借入金マスタ, w解約実行日)      ' 08/12/24 V189
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_借入金テーブル作成_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入金テーブル作成() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

'------------------------------------------------
' MBD010_借入金テーブルRead
'------------------------------------------------
Private Function MBD010_借入金テーブルRead(p借入金テーブル As MAA910_借入金テーブル) As MAA910_借入金テーブル
'
    Dim wFind As Boolean
    Dim w配列数 As Integer
    Dim j As Integer
'
    On Error GoTo MBD010_借入金テーブルRead_ERR
'
    w配列数 = UBound(G借入金テーブル)
    
    wFind = False
    For j = 1 To w配列数
        If p借入金テーブル.借入番号 = G借入金テーブル(j).借入番号 _
        And p借入金テーブル.実際年月日 = G借入金テーブル(j).実際年月日 _
        And p借入金テーブル.据置x回目 = G借入金テーブル(j).据置x回目 Then     ' 08/12/06 V189

            wFind = True
            Exit For
        End If
    Next
    
  '   And p借入金テーブル.返済回数 = G借入金テーブル(J).返済回数 _
  '      And p借入金テーブル.据置X回目 = G借入金テーブル(J).据置X回目 _


    If wFind Then
        MBD010_借入金テーブルRead.借入番号 = G借入金テーブル(j).借入番号
        MBD010_借入金テーブルRead.返済回数 = G借入金テーブル(j).返済回数
        MBD010_借入金テーブルRead.据置x回目 = G借入金テーブル(j).据置x回目
        MBD010_借入金テーブルRead.返済予定年月 = G借入金テーブル(j).返済予定年月
    
        MBD010_借入金テーブルRead.実際年月日 = G借入金テーブル(j).実際年月日
        MBD010_借入金テーブルRead.利息計算年月日 = G借入金テーブル(j).利息計算年月日    '10/01/04
        MBD010_借入金テーブルRead.返済金額 = G借入金テーブル(j).返済金額
        MBD010_借入金テーブルRead.元金額 = G借入金テーブル(j).元金額
        MBD010_借入金テーブルRead.利息額 = G借入金テーブル(j).利息額
        MBD010_借入金テーブルRead.保証料 = G借入金テーブル(j).保証料
        MBD010_借入金テーブルRead.手数料 = G借入金テーブル(j).手数料        ' 08/12/06 V189
        MBD010_借入金テーブルRead.金融保証料 = G借入金テーブル(j).金融保証料
        MBD010_借入金テーブルRead.融資残高 = G借入金テーブル(j).融資残高
        MBD010_借入金テーブルRead.日割日数 = G借入金テーブル(j).日割日数
        MBD010_借入金テーブルRead.利息対象期間日数 = G借入金テーブル(j).利息対象期間日数 'V182 2008/01/29
        MBD010_借入金テーブルRead.利率 = G借入金テーブル(j).利率
    Else
        MBD010_借入金テーブルRead.借入番号 = p借入金テーブル.借入番号
        MBD010_借入金テーブルRead.返済回数 = p借入金テーブル.返済回数
        MBD010_借入金テーブルRead.据置x回目 = p借入金テーブル.据置x回目
        MBD010_借入金テーブルRead.返済予定年月 = p借入金テーブル.返済予定年月
    
        MBD010_借入金テーブルRead.実際年月日 = p借入金テーブル.実際年月日       ' 08/12/14 V189
        MBD010_借入金テーブルRead.返済金額 = 0
        MBD010_借入金テーブルRead.元金額 = 0
        MBD010_借入金テーブルRead.利息額 = 0
        MBD010_借入金テーブルRead.保証料 = 0
        MBD010_借入金テーブルRead.手数料 = 0                    ' 08/12/06 V189
        MBD010_借入金テーブルRead.金融保証料 = 0
        MBD010_借入金テーブルRead.融資残高 = 0
        MBD010_借入金テーブルRead.日割日数 = 0
        MBD010_借入金テーブルRead.利息対象期間日数 = 0 'V182 2008/01/29
        MBD010_借入金テーブルRead.利率 = 0
    End If
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_借入金テーブルRead_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入金テーブルRead() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

'------------------------------------------------
' MBD010_借入金テーブルWrite
'------------------------------------------------
Private Sub MBD010_借入金テーブルWrite(p借入金テーブル As MAA910_借入金テーブル, p借入金計画 As MAA910_借入金) '2016/09/23
'
    Dim wFind As Boolean
    Dim w配列数 As Integer
    Dim j As Integer
    
    Dim wk日数 As Integer       '2016/09/23
    
'
    On Error GoTo MBD010_借入金テーブルWrite_ERR
    
  'If p借入金テーブル.元金額 <> 0 Or p借入金テーブル.利息額 <> 0 Or p借入金テーブル.保証料 <> 0 _
  '   Or p借入金テーブル.手数料 Then         '10/02/27
'
    w配列数 = UBound(G借入金テーブル)
    
    wFind = False
    For j = 1 To w配列数
        If p借入金テーブル.借入番号 = G借入金テーブル(j).借入番号 _
        And p借入金テーブル.実際年月日 = G借入金テーブル(j).実際年月日 _
        And p借入金テーブル.据置x回目 = G借入金テーブル(j).据置x回目 Then     ' 08/12/06 V189

            wFind = True
            Exit For
        End If
    Next
    
    '   And p借入金テーブル.返済回数 = G借入金テーブル(J).返済回数 _
    '    And p借入金テーブル.据置X回目 = G借入金テーブル(J).据置X回目 _


    If Not wFind Then
        j = w配列数 + 1
        ReDim Preserve G借入金テーブル(j)
    End If
    
    
    
    
    '***** 中間利払最終日控除の対応　　2016/09/23
    If p借入金計画.利息控除区分 = 4 Then                    '2016/09/23
        If p借入金計画.利息区分 = XMXA020_区分("利息区分", "利息先払") Then         '2016/809/23
            '***利息先払
            '**実行日
            If p借入金計画.実行日 = p借入金テーブル.実際年月日 Then                 '2016/09/23
                p借入金テーブル.日割日数 = p借入金テーブル.日割日数 - 1             '2016/09/23
                p借入金テーブル.利息対象期間日数 = p借入金テーブル.利息対象期間日数 - 1 '2016/09/23
                p借入金テーブル.利息額 = MBD010_利息計算小数点5桁(p借入金テーブル.利率, p借入金テーブル.融資残高, _
                        p借入金テーブル.日割日数, p借入金計画.金利計算年間日数)     '2016/09/23
            End If                                                                  '2016/09/23
            
        Else                                                                        '2016/09/23
            '***利息後払
            '**実行日
            If j > 1 Then                                                           '2016/09/23
                wk日数 = DateDiff("d", p借入金計画.実行日, p借入金テーブル.利息計算年月日) + 1  '2016/09/23
                
                If wk日数 = p借入金テーブル.日割日数 Then        '2016/09/23
                    '**実行日から、次回支払日の前日までの調整
                    p借入金テーブル.日割日数 = p借入金テーブル.日割日数 - 1         '2016/09/23
                    p借入金テーブル.利息対象期間日数 = p借入金テーブル.利息対象期間日数 - 1     '2016/09/23
                    p借入金テーブル.利息額 = MBD010_利息計算小数点5桁(p借入金テーブル.利率, p借入金テーブル.融資残高 + p借入金テーブル.元金額, _
                        p借入金テーブル.日割日数, p借入金計画.金利計算年間日数)     '2016/09/23
                End If                                                              '2016/09/23
            End If                                                                  '2016/09/23
        End If                                                                      '2016/09/23
    End If                                                                          '2016/09/23
    
                
                
    
    
    
    
    

    '** テーブルにセット **
    G借入金テーブル(j).借入番号 = p借入金テーブル.借入番号
    G借入金テーブル(j).返済回数 = p借入金テーブル.返済回数
    G借入金テーブル(j).据置x回目 = p借入金テーブル.据置x回目
    G借入金テーブル(j).返済予定年月 = p借入金テーブル.返済予定年月
    
    G借入金テーブル(j).実際年月日 = p借入金テーブル.実際年月日
    G借入金テーブル(j).利息計算年月日 = p借入金テーブル.利息計算年月日  '10/01/04
    G借入金テーブル(j).返済金額 = p借入金テーブル.返済金額
    G借入金テーブル(j).元金額 = p借入金テーブル.元金額
    G借入金テーブル(j).利息額 = p借入金テーブル.利息額
    G借入金テーブル(j).保証料 = p借入金テーブル.保証料
    G借入金テーブル(j).手数料 = p借入金テーブル.手数料          ' 08/12/06 V189
    G借入金テーブル(j).金融保証料 = p借入金テーブル.金融保証料
    G借入金テーブル(j).融資残高 = p借入金テーブル.融資残高
    G借入金テーブル(j).日割日数 = p借入金テーブル.日割日数
    G借入金テーブル(j).利息対象期間日数 = p借入金テーブル.利息対象期間日数 'V182 2008/01/29
    G借入金テーブル(j).利率 = p借入金テーブル.利率
    
  'End If                    '10/02/27
  
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_借入金テーブルWrite_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入金テーブルWrite() でエラー" + vbCrLf + vbCrLf + _
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
' MBD010_金利参照
'------------------------------------------------
Private Function MBD010_金利参照(p借入金 As MAA910_借入金, p対象年月 As Date) As Double
'
    Dim j As Integer
'
    On Error GoTo MBD010_金利参照_ERR
'
    MBD010_金利参照 = p借入金.利率
    
    '***未来利息シュミュレーション
    If G金利SM = True And p借入金.金利種別 = 0 Then         '10/11/11 V189R
        For j = 2 To 200                                    '10/11/11 V189R
            If IsNull(wSM年月日(j)) Then                    '10/11/11 V189R
                Exit For                                    '10/11/11 V189R
            End If                                          '10/11/11 V189R
            
            If p対象年月 >= wSM年月日(j) Then               '10/11/11 V189R
                MBD010_金利参照 = wSM利率(j)                '10/11/11 V189R
            Else                                            '10/11/11 V189R
                Exit For                                    '10/11/11 V189R
            End If                                          '10/11/11 V189R
        Next                                                '10/11/11 V189R
        
    Else                                                    '10/11/11 V189R
        
        For j = 2 To 100                                        ' 08/07/04 V188
            If IsNull(p借入金.金利(j).金利変更x回目年月) Then
                Exit For
            End If
        
             
            If p対象年月 >= p借入金.金利(j).金利変更x回目年月 Then
                MBD010_金利参照 = p借入金.金利(j).金利x回目
            Else
                Exit For
            End If
        Next
    End If                                                  '10/11/11 V189R
    
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_金利参照_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_金利参照() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

'------------------------------------------------
' MBD010_保証料算出
'------------------------------------------------
Private Function MBD010_保証料算出(p借入番号 As String, _
                          p自己資金 As Integer, _
                          p有担保 As Integer, _
                          p保証料分割 As Integer, _
                          p保証料率 As Double, _
                          p融資金額 As Double, _
                          p実行年月日 As Variant, _
                          p解約実行日 As Variant, _
                          p初回返済実行日 As Variant, _
                          p最終返済実行日 As Variant, _
                          p初回返済年月 As Variant, _
                          p最終返済年月 As Variant, _
                          pw支払回数 As Long, _
                          pw据置回数 As Long, _
                          p実行支払年月 As Variant, _
                          p営業日区分 As Integer _
                          ) As MBD010_保証料算出リターン
'
    Dim wDbl1 As Double
    Dim j As Integer
    
    Dim w解約保証料戻 As Double
    
    Dim w据置期間部分_金額 As Double    '据置期間部分
    Dim w均等4分割部分_金額 As Double   '均等４分割部分
    Dim w合算_金額 As Double            '据置期間部分 + 均等４分割部分
    
    Dim w分割 As Integer   '1...分割  0...一括
    
    Dim w初回保証料 As Double
    Dim w保証料X年後(9) As Double 'X年後
    
    Dim w実行年 As Integer
    Dim w実行月 As Integer
    
    Dim w返済予定年月 As Date
    Dim w実際予定年月日 As Date
    Dim w解約実行日 As Date    '自己資金
    
    
    Dim w初回回数 As Long
    Dim w解約回数 As Long
    
    Dim w保証期間 As Long
       
      
    '解約保証料戻
    Dim w経過期間a As Long
    Dim w経過期間b As Long
    Dim w経過期間c As Long
    Dim w金額1 As Double
    Dim w金額2 As Double
    Dim w金額3 As Double
           
    Dim w経過期間B_A As Long
    Dim w経過期間C_B As Long
    Dim w経過期間C_A As Long
    Dim w経過期間B_C As Long
        
    Dim w保証_経過XA As Double
    Dim w保証_経過XB As Double
    Dim w保証_経過XC As Double
'
    On Error GoTo MBD010_保証料算出_ERR
'
    ' =========================================
    '                   前処理
    ' =========================================
    '** 一括保証料計算 **
 '   w保証期間 = pw支払回数 + pw据置回数
    pw支払回数 = DateDiff("m", p初回返済年月, p最終返済年月) + 1
        
    'w保証期間 = DateDiff("m", C年月日.GetDate("月始", p実行支払年月), _
    '                              C年月日.GetDate("月始", p最終返済実行日))
    w保証期間 = pw支払回数 + pw据置回数
    'If Day(p最終返済実行日) >= Day(p実行年月日) Then
    '    w保証期間 = w保証期間 + 1
    'End If
        
       
    
    wDbl1 = (pw据置回数 * p保証料率 / 12) / 100        '2004/3/27
    w据置期間部分_金額 = P8.FRound(p融資金額 * wDbl1, 0)
    
    wDbl1 = (pw支払回数 * p保証料率 * MBD010_保証料分割係数(w保証期間) / 12) / 100
    w均等4分割部分_金額 = p融資金額 * wDbl1
    w均等4分割部分_金額 = P8.FRound(w均等4分割部分_金額, 0)
    
    
    
    w据置期間部分_金額 = P8.FRound(w据置期間部分_金額, 1)
  '  w据置期間部分_金額 = Fix((w据置期間部分_金額 + 5) / 10)
  '  w据置期間部分_金額 = w据置期間部分_金額 * 10
    
    w均等4分割部分_金額 = P8.FRound(w均等4分割部分_金額, 1)
  '  w均等4分割部分_金額 = Fix((w均等4分割部分_金額 + 5) / 10)
  '  w均等4分割部分_金額 = w均等4分割部分_金額 * 10
    
    w合算_金額 = w据置期間部分_金額 + w均等4分割部分_金額
    
    '** 分割保証料計算 **
    w分割 = 0
    w初回保証料 = w合算_金額
    
    For j = 1 To 9
        w保証料X年後(j) = 0
    Next

    w実行年 = Year(p実行年月日)
    w実行月 = Month(p実行年月日)
    
    If IsNull(p初回返済実行日) Then
        w初回回数 = 0
    Else
        w初回回数 = DateDiff("m", p実行年月日, p初回返済実行日)
    End If
    
    If IsDate(p解約実行日) Then
        GDate1 = p解約実行日
        If p自己資金 = 1 Then
            GDate1 = DateAdd("d", -15, GDate1)
        End If
        w解約回数 = DateDiff("m", p実行年月日, GDate1)
    End If
    
    ' =========================================
    '       　       通常 保証料計算
    ' =========================================
    'If p有担保 = 1 _
    'And (p融資金額 > 15000000 Or p融資金額 < -15000000)
    If w保証期間 > 24 _
    And p保証料分割 = 1 Then
         
        w分割 = 1
         
        Select Case w保証期間
        
            Case Is <= 48      '３－４年
                w初回保証料 = P8.FRound(Fix((75 * w合算_金額) / 100), 1)
                
                wDbl1 = w合算_金額 - w初回保証料
                w保証料X年後(1) = P8.FRound(wDbl1, 1)
         
            
            Case Is <= 72       '５－６年
                w初回保証料 = P8.FRound(Fix((60 * w合算_金額) / 100), 1)
                w保証料X年後(1) = P8.FRound(Fix((30 * w合算_金額) / 100), 1)
               
                wDbl1 = w合算_金額 - w初回保証料 - w保証料X年後(1)
                w保証料X年後(2) = P8.FRound(wDbl1, 1)
         
            Case Is <= 96      '７－８年
                w初回保証料 = P8.FRound(Fix((45 * w合算_金額) / 100), 1)
                w保証料X年後(1) = P8.FRound(Fix((35 * w合算_金額) / 100), 1)
                w保証料X年後(2) = P8.FRound(Fix((15 * w合算_金額) / 100), 1)
               
                wDbl1 = w合算_金額 - w初回保証料 - w保証料X年後(1) - w保証料X年後(2)
                w保証料X年後(3) = P8.FRound(wDbl1, 1)
                     
            Case Is <= 120        '９－１０年
                w初回保証料 = P8.FRound(Fix((35 * w合算_金額) / 100), 1)
                w保証料X年後(1) = P8.FRound(Fix((30 * w合算_金額) / 100), 1)
                w保証料X年後(2) = P8.FRound(Fix((20 * w合算_金額) / 100), 1)
                w保証料X年後(3) = P8.FRound(Fix((10 * w合算_金額) / 100), 1)
      
                wDbl1 = w合算_金額 - w初回保証料 - w保証料X年後(1) - w保証料X年後(2) - w保証料X年後(3)
                w保証料X年後(4) = P8.FRound(wDbl1, 1)
                
            Case Is <= 144        '１１－１２年
                w初回保証料 = P8.FRound(Fix((30 * w合算_金額) / 100), 1)
                w保証料X年後(1) = P8.FRound(Fix((20 * w合算_金額) / 100), 1)
                w保証料X年後(2) = P8.FRound(Fix((20 * w合算_金額) / 100), 1)
                w保証料X年後(3) = P8.FRound(Fix((15 * w合算_金額) / 100), 1)
                w保証料X年後(4) = P8.FRound(Fix((10 * w合算_金額) / 100), 1)
      
                wDbl1 = w合算_金額 - w初回保証料 - w保証料X年後(1) - w保証料X年後(2) - w保証料X年後(3)
                wDbl1 = wDbl1 - w保証料X年後(4)
                w保証料X年後(5) = P8.FRound(wDbl1, 1)
                
            Case Is <= 168       '１３－１４年
                w初回保証料 = P8.FRound(Fix((25 * w合算_金額) / 100), 1)
                w保証料X年後(1) = P8.FRound(Fix((20 * w合算_金額) / 100), 1)
                w保証料X年後(2) = P8.FRound(Fix((20 * w合算_金額) / 100), 1)
                w保証料X年後(3) = P8.FRound(Fix((15 * w合算_金額) / 100), 1)
                w保証料X年後(4) = P8.FRound(Fix((10 * w合算_金額) / 100), 1)
                w保証料X年後(5) = P8.FRound(Fix((5 * w合算_金額) / 100), 1)
      
                wDbl1 = w合算_金額 - w初回保証料 - w保証料X年後(1) - w保証料X年後(2) - w保証料X年後(3)
                wDbl1 = wDbl1 - w保証料X年後(4) - w保証料X年後(5)
                w保証料X年後(6) = P8.FRound(wDbl1, 1)
                
            Case Is <= 192       '１５－１６年
                w初回保証料 = P8.FRound(Fix((20 * w合算_金額) / 100), 1)
                w保証料X年後(1) = P8.FRound(Fix((20 * w合算_金額) / 100), 1)
                w保証料X年後(2) = P8.FRound(Fix((15 * w合算_金額) / 100), 1)
                w保証料X年後(3) = P8.FRound(Fix((15 * w合算_金額) / 100), 1)
                w保証料X年後(4) = P8.FRound(Fix((10 * w合算_金額) / 100), 1)
                w保証料X年後(5) = P8.FRound(Fix((10 * w合算_金額) / 100), 1)
                w保証料X年後(6) = P8.FRound(Fix((5 * w合算_金額) / 100), 1)
      
                wDbl1 = w合算_金額 - w初回保証料 - w保証料X年後(1) - w保証料X年後(2) - w保証料X年後(3)
                wDbl1 = wDbl1 - w保証料X年後(4) - w保証料X年後(5) - w保証料X年後(6)
                w保証料X年後(7) = P8.FRound(wDbl1, 1)
                
             Case Is <= 216      '１７－１８年
                w初回保証料 = P8.FRound(Fix((20 * w合算_金額) / 100), 1)
                w保証料X年後(1) = P8.FRound(Fix((20 * w合算_金額) / 100), 1)
                w保証料X年後(2) = P8.FRound(Fix((15 * w合算_金額) / 100), 1)
                w保証料X年後(3) = P8.FRound(Fix((15 * w合算_金額) / 100), 1)
                w保証料X年後(4) = P8.FRound(Fix((10 * w合算_金額) / 100), 1)
                w保証料X年後(5) = P8.FRound(Fix((5 * w合算_金額) / 100), 1)
                w保証料X年後(6) = P8.FRound(Fix((5 * w合算_金額) / 100), 1)
                w保証料X年後(7) = P8.FRound(Fix((5 * w合算_金額) / 100), 1)
      
                wDbl1 = w合算_金額 - w初回保証料 - w保証料X年後(1) - w保証料X年後(2) - w保証料X年後(3)
                wDbl1 = wDbl1 - w保証料X年後(4) - w保証料X年後(5) - w保証料X年後(6) - w保証料X年後(7)
                w保証料X年後(8) = P8.FRound(wDbl1, 1)
                 
                  
                 
                  
                 
                    
            Case Is <= 240        '１９－２０年
                w初回保証料 = P8.FRound(Fix((20 * w合算_金額) / 100), 1)
                w保証料X年後(1) = P8.FRound(Fix((20 * w合算_金額) / 100), 1)
                w保証料X年後(2) = P8.FRound(Fix((15 * w合算_金額) / 100), 1)
                w保証料X年後(3) = P8.FRound(Fix((15 * w合算_金額) / 100), 1)
                w保証料X年後(4) = P8.FRound(Fix((10 * w合算_金額) / 100), 1)
                w保証料X年後(5) = P8.FRound(Fix((5 * w合算_金額) / 100), 1)
                w保証料X年後(6) = P8.FRound(Fix((5 * w合算_金額) / 100), 1)
                w保証料X年後(7) = P8.FRound(Fix((5 * w合算_金額) / 100), 1)
                w保証料X年後(8) = P8.FRound(Fix((3 * w合算_金額) / 100), 1)
             
                wDbl1 = w合算_金額 - w初回保証料 - w保証料X年後(1) - w保証料X年後(2)
                wDbl1 = wDbl1 - w保証料X年後(3) - w保証料X年後(4) - w保証料X年後(5)
                wDbl1 = wDbl1 - w保証料X年後(6) - w保証料X年後(7) - w保証料X年後(8)
                w保証料X年後(9) = P8.FRound(wDbl1, 1)
          
        End Select
    End If
      
    ' =========================================
    '       　      解約時 保証料計算
    ' =========================================
    w解約保証料戻 = 0
     
    If IsDate(p解約実行日) Then
 
        w経過期間c = DateDiff("m", C年月日.GetDate("月始", p実行支払年月), _
                                  C年月日.GetDate("月始", p初回返済実行日))
        If Day(p初回返済実行日) >= Day(p実行年月日) Then
            w経過期間c = w経過期間c + 1
        End If
        
        w解約実行日 = p解約実行日
        If p自己資金 = 1 Then
            w解約実行日 = DateDiff("d", 15, w解約実行日)
        End If
        
        w経過期間a = DateDiff("m", C年月日.GetDate("月始", p実行年月日), _
                                  C年月日.GetDate("月始", w解約実行日))
        If Day(w解約実行日) >= Day(p実行年月日) Then
            w経過期間a = w経過期間a + 1
        End If
        
        w経過期間b = Fix((w経過期間a - 1) / 12)
        w経過期間b = (w経過期間b + 1) * 12
        
        If w経過期間b > w保証期間 Then
            w経過期間b = w保証期間
        End If
        
        w経過期間B_A = w経過期間b - w経過期間a
        w経過期間C_B = w経過期間c - w経過期間b
        w経過期間C_A = w経過期間c - w経過期間a
        w経過期間B_C = w経過期間b - w経過期間c
        
        If pw支払回数 <> 0 Then
            w保証_経過XA = (w保証期間 - w経過期間a) / pw支払回数
            w保証_経過XB = (w保証期間 - w経過期間b) / pw支払回数
            w保証_経過XC = (w保証期間 - w経過期間c) / pw支払回数
        Else
            w保証_経過XA = 0
            w保証_経過XB = 0
            w保証_経過XC = 0
        End If
        
            
        
        w解約保証料戻 = 0
        
        '** 初回返済実行日以前の解約その１ **
        If 0 <= w経過期間B_A And 0 < w経過期間C_B Then
            wDbl1 = (w経過期間B_A) * p保証料率 / 1200
            w金額1 = Round((p融資金額) * wDbl1 * 0.9)
    
            wDbl1 = (w経過期間C_B - 1) * p保証料率 / 1200
            w金額2 = Round((p融資金額) * wDbl1)
                    
            w解約保証料戻 = P8.FRound((w金額1 + w金額2 + w均等4分割部分_金額), 1) * -1
        End If
           
        '** 初回返済実行日以前の解約その２ **
        If 0 < w経過期間C_A And 0 <= w経過期間B_C Then
            wDbl1 = (w経過期間C_A - 1) * p保証料率 / 1200
            w金額1 = Round((p融資金額) * wDbl1 * 0.9)
                                     
            w保証_経過XC = (w保証期間 - pw据置回数) / pw支払回数
            w金額3 = (w保証_経過XC) * (w保証_経過XC)
            w金額3 = w金額3 - (w保証_経過XB) * (w保証_経過XB)    '2004/3/26
            w金額3 = Round(w均等4分割部分_金額 * w金額3 * 0.9)
                    
            w金額2 = (w保証_経過XB) * (w保証_経過XB)
            w金額2 = Round(w均等4分割部分_金額 * w金額2)
                    
            w解約保証料戻 = P8.FRound((w金額1 + w金額2 + w金額3), 1) * -1
        End If
        
             
        '** 初回返済実行日後の解約 **
        If w経過期間a >= w経過期間c Then
            w金額1 = (w保証_経過XA) * (w保証_経過XA)
            w金額1 = w金額1 - (w保証_経過XB) * (w保証_経過XB)
            w金額1 = Round(w均等4分割部分_金額 * w金額1 * 0.9)
                    
            w金額2 = (w保証_経過XB) * (w保証_経過XB)
            w金額2 = Round(w均等4分割部分_金額 * w金額2)
                    
       '     w解約保証料戻 = Round((w金額1 + w金額2)) * -1
            w解約保証料戻 = P8.FRound((w金額1 + w金額2), 1) * -1
       '     w解約保証料戻 = (w金額1 + w金額2)
       '     w解約保証料戻 = w解約保証料戻 + 5
       '     w解約保証料戻 = Fix(w解約保証料戻 / 10)
       '     w解約保証料戻 = w解約保証料戻 * 10 * -1
        End If

        '** 分割時計算 **
        If w分割 = 1 Then
            w解約保証料戻 = w解約保証料戻 + w合算_金額 - w初回保証料
            For j = 1 To 9
                GDate1 = DateSerial(w実行年 + j, w実行月, 1)  '2003/07/04
                GDate1 = MXA030_翌営業年月日計算(GDate1, w保証料支払日, p営業日区分) ' 07/01/30 V180
                If p解約実行日 > GDate1 Then
                    w解約保証料戻 = w解約保証料戻 - w保証料X年後(j)
                End If
            Next
        End If
    End If
    
    MBD010_保証料算出.解約保証料戻 = w解約保証料戻
    MBD010_保証料算出.初回保証料 = w初回保証料
    For j = 1 To 9
        MBD010_保証料算出.保証料X年後(j) = w保証料X年後(j)
    Next
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_保証料算出_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_保証料算出() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

'------------------------------------------------
' MBD010_保証料分割係数
'------------------------------------------------
Private Function MBD010_保証料分割係数(p保証期間 As Long) As Double
'
    Select Case p保証期間
        Case Is <= 6:  MBD010_保証料分割係数 = 0.7
        Case Is <= 12: MBD010_保証料分割係数 = 0.65
        Case Is <= 24: MBD010_保証料分割係数 = 0.6
        Case Else:     MBD010_保証料分割係数 = 0.55
    End Select
'
End Function

'------------------------------------------------
' MBD010_内入処理
'------------------------------------------------
Public Sub MBD010_内入処理(p借入金マスタ As MAA910_借入金)
'
    Dim pTable As MAA910_借入金テーブル
    Dim k As Integer
    Dim w配列数 As Integer      '10/01/30
    Dim j As Integer            '10/01/30
    
    
    Dim w対象年月 As Date
    Dim p対象年月 As Date
    Dim w対象年月日 As Date
    Dim w年 As Integer
    Dim w月 As Integer
    Dim w日 As Integer
    
    Dim w実行日加算 As Integer
    Dim w利息計算基準年月日 As Date     ' 08/12/10 V189
    
    Dim w調整内入開始年月日 As Date     '10/02/07
    Dim w調整内入終了年月日 As Date     '10/02/07
'
    On Error GoTo MBD010_内入処理_ERR
'
    '***ウチイレ処理無視
    If w解約無効F = 1 Then                      '10/01/30
        Exit Sub                                '10/01/30
    End If                                      '10/01/30
'
    If Format(p借入金マスタ.実行日, "yyyy/mm/dd") = Format(w内入開始年月日, "yyyy/mm/dd") _
        And p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息後払") Then
        w実行日加算 = 1
    Else
        w実行日加算 = 0
    End If
    
    If IsNull(w内入開始年月日) Then
        Exit Sub
    End If

    '内入明細
    'If w借入内入.内入区分 <> True Then
    '    Exit Sub
    'End If
'
    For k = 1 To 500 '内入回数
        If IsNull(w借入内入.内入(k).内入x回目年月日) Or P8.FCStr(w借入内入.内入(k).内入x回目年月日) = "" Then
            Exit For
        End If
        
        If Format(p借入金マスタ.最終返済実行日, "yyyy/mm/dd") _
            <> Format(w借入内入.内入(k).内入x回目年月日, "yyyy/mm/dd") Then ' 09/01/21 V189

        
            If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
                If Format(w内入開始年月日, "yyyy/mm/dd") > Format(w借入内入.内入(k).内入x回目年月日, "yyyy/mm/dd") Then
                    GoTo next1
                End If
        
                If Format(w内入終了年月日, "yyyy/mm/dd") <= Format(w借入内入.内入(k).内入x回目年月日, "yyyy/mm/dd") Then
                    GoTo next1
                End If
            End If
        
            If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息後払") Then
                If Format(w内入開始年月日, "yyyy/mm/dd") >= Format(w借入内入.内入(k).内入x回目年月日, "yyyy/mm/dd") Then
                    GoTo next1
                End If
        
                If Format(w内入終了年月日, "yyyy/mm/dd") < Format(w借入内入.内入(k).内入x回目年月日, "yyyy/mm/dd") Then
                    GoTo next1
                End If
            End If
        End If                          ' 09/01/21 V189
        
        
        '**標準の利息計算年月を参照
        w利息計算基準年月日 = w借入内入.内入(k).内入x回目年月日
        pTable.利息計算年月日 = w借入内入.内入(k).内入x回目年月日       '10/01/30
        'w配列数 = UBound(G借入金テーブル)           '10/01/30
        'For j = 1 To w配列数                        '10/01/30
        '    If Format(w利息計算基準年月日, "yyyy/mm/dd") = Format(G借入金テーブル(j).実際年月日, "yyyy/mm/dd") _
        '        And G借入金テーブル(j).据置X回目 = 2 Then                   '10/01/30
        '        w利息計算基準年月日 = G借入金テーブル(j).利息計算年月日     '10/01/30
        '        pTable.利息計算年月日 = G借入金テーブル(j).利息計算年月日   '10/01/30
        '        Exit For                                                    '10/01/30
        '    End If                                                          '10/01/30
        'Next                                                                '10/01/30
        
        'wv01 = MXA030_翌営業年月日計算(CDate(wv01), p借入金マスタ.支払日, p借入金マスタ.営業日区分)  '10/01/20
        
        
        
        w年 = Year(w利息計算基準年月日)         '10/02/08
        w月 = Month(w利息計算基準年月日)        '10/02/08
        p対象年月 = Right("0000" & CStr(w年), 4) & "/" & Right("00" & CStr(w月), 2) & "/" & Right("00" & CStr(1), 2)
        w対象年月 = p対象年月
        w対象年月日 = MXA030_翌営業年月日計算(CDate(w対象年月), p借入金マスタ.支払日, p借入金マスタ.営業日区分)
        If Format(w対象年月日, "yyyy/mm/dd") = Format(w利息計算基準年月日, "yyyy/mm/dd") Then '10/02/08
            w利息計算基準年月日 = MBD010_利息計算年月日(w対象年月, _
                        p借入金マスタ.支払日, p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)
            
        Else
            w対象年月 = DateAdd("m", -1, p対象年月)
            w対象年月日 = MXA030_翌営業年月日計算(CDate(w対象年月), _
                                    p借入金マスタ.支払日, p借入金マスタ.営業日区分)
            If Format(w対象年月日, "yyyy/mm/dd") = Format(w利息計算基準年月日, "yyyy/mm/dd") Then '10/02/08
                w利息計算基準年月日 = MBD010_利息計算年月日(w対象年月, _
                        p借入金マスタ.支払日, p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)
                
            Else
                w対象年月 = DateAdd("m", 1, p対象年月)
                w対象年月日 = MXA030_翌営業年月日計算(CDate(DateAdd("m", 1, w対象年月)), _
                                    p借入金マスタ.支払日, p借入金マスタ.営業日区分)
                If Format(w対象年月日, "yyyy/mm/dd") = Format(w利息計算基準年月日, "yyyy/mm/dd") Then '10/02/08
                    w利息計算基準年月日 = MBD010_利息計算年月日(w対象年月, _
                        p借入金マスタ.支払日, p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)
                    
                End If
            End If
        End If
        
        pTable.利息計算年月日 = w利息計算基準年月日 '10/02/25
            
            
            
        If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
            If Format(w借入内入.内入(k).内入x回目年月日, "yyyy/mm/dd") = Format(p借入金マスタ.最終返済実行日, "yyyy/mm/dd") Then ' 0903/01 V189
                pTable.日割日数 = 0 ' 09/03/01 V189
                pTable.利息対象期間日数 = 0 ' 09/03/01 V189
            Else
                '* 利息先払　最終支払期間中の内入終了年月日調整
                '* 初回返済実行日の調整 10/02/09
                If Format(w内入終了年月日, "yyyy/mm/dd") = Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then
                    w調整内入終了年月日 = MBD010_利息計算年月日(p借入金マスタ.初回返済実行日, _
                        p借入金マスタ.支払日, p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)  '10/02/07
                    If Format(GDate利息対象年月日, "yyyy/mm/dd") <> Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") Then '120/02/07
                        w調整内入終了年月日 = p借入金マスタ.初回返済実行日          '10/02/07
                    Else                                                            '10/02/07
                        w調整内入終了年月日 = w内入終了年月日                       '10/02/07
                    End If                                                          '10/02/07
                Else                                                                '10/02/07
                    w調整内入終了年月日 = w内入終了年月日                           '10/02/07
                End If                                                              '10/02/07
                
                
                If Format(w調整内入終了年月日, "yyyy/mm/dd") _
                        = Format(p借入金マスタ.初回返済実行日, "yyyy/mm/dd") _
                   And (p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3) Then  '10/02/07
                   w調整内入終了年月日 = w調整内入終了年月日 - 1                '10/02/07
                End If
                
                
                
                '* 最終返済実行日の調整
                If Format(w内入終了年月日, "yyyy/mm/dd") = Format(p借入金マスタ.最終返済実行日, "yyyy/mm/dd") Then
                    w調整内入終了年月日 = MBD010_利息計算年月日(p借入金マスタ.最終返済年月, _
                        p借入金マスタ.支払日, p借入金マスタ.営業日区分, p借入金マスタ.利息計算日数区分)  '10/02/07
                    If Format(GDate利息対象年月日, "yyyy/mm/dd") <> Format(p借入金マスタ.最終返済実行日, "yyyy/mm/dd") Then '120/02/07
                        w調整内入終了年月日 = p借入金マスタ.最終返済実行日          '10/02/07
                    Else                                                            '10/02/07
                        w調整内入終了年月日 = w内入終了年月日                       '10/02/07
                    End If                                                          '10/02/07
                Else                                                                '10/02/07
                    w調整内入終了年月日 = w内入終了年月日                           '10/02/07
                End If                                                              '10/02/07
                
                
                If Format(w調整内入終了年月日, "yyyy/mm/dd") _
                        = Format(p借入金マスタ.最終返済実行日, "yyyy/mm/dd") _
                   And (p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3) Then  '10/02/07
                   w調整内入終了年月日 = w調整内入終了年月日 - 1                '10/02/07
                End If                                                          '10/02/07
                
                
                pTable.日割日数 = (DateDiff("d", w利息計算基準年月日, w調整内入終了年月日) + w実行日加算) * -1  '10/02/07
                pTable.利息対象期間日数 = pTable.日割日数                   '10/01/30
                
            End If                  ' 09/03/01 V189
            
        Else
            '* 利息後払　金利初回年月までの期間の内入開始年月日の調整
            If Format(p借入金マスタ.実行日, "yyyy/mm/dd") _
                = Format(w内入開始年月日, "yyyy/mm/dd") _
                And (p借入金マスタ.利息控除区分 = 2 Or p借入金マスタ.利息控除区分 = 3) Then
                w調整内入開始年月日 = DateAdd("d", 1, p借入金マスタ.実行日)         '10/02/07
            Else                                                                    '10/02/07
                w調整内入開始年月日 = w内入開始年月日                               '10/02/07
            End If                                                                  '10/02/07
            
            pTable.日割日数 = DateDiff("d", w調整内入開始年月日, w利息計算基準年月日) + w実行日加算
            pTable.利息対象期間日数 = pTable.日割日数                       '10/01/30
        End If
        
        w対象年月 = w借入内入.内入(k).内入x回目年月日
        'pTable.利率 = MBD010_金利参照(p借入金マスタ, w対象年月)
        
        '*** 変動金利　設定
        If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then   ' 08/12/10 V189
            If Not IsNull(w解約実行日) And _
                    Format(pTable.実際年月日, "yyyy/mm/dd") = Format(w解約実行日, "yyyy/mm/dd") Then
                    w利息計算基準年月日 = DateAdd("d", -1, w借入内入.内入(k).内入x回目年月日)   '10/01/23
            Else                                                                '10/01/23
                    'w利息計算基準年月日 = w借入内入.内入(K).内入x回目年月日                     '10/01/23
            End If                                                              '10/01/23
            'w利息計算基準年月日 = w借入内入.内入(K).内入x回目年月日     ' 08/12/10 V189
        Else                                                                ' 08/12/10 V189
            w利息計算基準年月日 = DateAdd("d", -pTable.利息対象期間日数, w利息計算基準年月日) ' 10/01/30 V189
        End If                                                              ' 08/12/10 V189
        pTable.利率 = MBD010_金利参照(p借入金マスタ, w利息計算基準年月日)   ' 08/12/10 V189
        
        pTable.利息額 = MBD010_利息計算小数点5桁(pTable.利率, _
                                                               w借入内入.内入(k).内入金額x回目, _
                              pTable.日割日数, p借入金マスタ.金利計算年間日数) '10/01/15
           
        pTable.借入番号 = p借入金マスタ.借入番号
        pTable.返済回数 = 0
        If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then       ' 08/12/16 V189
            pTable.据置x回目 = 3                    ' 08/12/06 V189
        Else                                        ' 08/12/06 V189
            pTable.据置x回目 = 1                    ' 08/12/06 V189
        End If                                      ' 08/12/06 V189
        'pTable.据置X回目 = 1                         '10/01/30
        
        
        
        pTable.実際年月日 = w借入内入.内入(k).内入x回目年月日
        
        w年 = Year(w借入内入.内入(k).内入x回目年月日)       ' 08/12/09 V189
        w月 = Month(w借入内入.内入(k).内入x回目年月日)      ' 08/12/09 V189
        w日 = Day(w借入内入.内入(k).内入x回目年月日)        ' 08/12/09 V189
        w対象年月 = CDate(CStr(w年) + "/" + CStr(w月) + "/" + CStr(1))  ' 08/12/09 V189
        w対象年月日 = MXA030_翌営業年月日計算(CDate(w対象年月), p借入金マスタ.支払日, p借入金マスタ.営業日区分) ' 08/12/09 V189
        If Format(w対象年月日, "yyyy/mm/dd") >= Format(pTable.実際年月日, "yyyy/mm/dd") Then ' 08/12/09 V189
            pTable.返済予定年月 = w対象年月                     ' 08/12/09 V189
        Else                                                    ' 08/12/09 V189
            pTable.返済予定年月 = DateAdd("m", 1, w対象年月)    ' 08/12/09 V189
        End If
        
                                                                    
       
        pTable.元金額 = w借入内入.内入(k).内入金額x回目
        pTable.返済金額 = pTable.元金額 + pTable.利息額
        w融資残高 = w融資残高 - pTable.元金額
        pTable.融資残高 = w融資残高
        pTable.保証料 = 0
        pTable.手数料 = w借入内入.内入(k).手数料x回目
        Call MBD010_借入金テーブルWrite(pTable, p借入金マスタ)              '2016/09/23
next1:
    Next
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_内入処理_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_内入処理() でエラー" + vbCrLf + vbCrLf + _
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
' MBD010_融資残高再計算
'------------------------------------------------
Public Sub MBD010_融資残高再計算(p借入金マスタ As MAA910_借入金, p解約実行日 As Variant)
'
    Dim w配列数 As Integer
    Dim j As Integer
    Dim X As Integer
    Dim p借入金テーブル As MAA910_借入金テーブル
    Dim p融資残高 As Double
    Dim w利息計算基準年月日 As Date '10/01/04
'
    On Error GoTo MBD010_融資残高再計算_ERR
    
    '** ウチイレ処理無視
    If w解約無効F = 1 Then              '10/01/30
        Exit Sub                        '10/01/30
    End If                              '10/01/30

    '内入明細
    'If w借入内入.内入区分 <> True Then
    '    Exit Sub
    'End If
'
    w配列数 = UBound(G借入金テーブル)
     
    X = 1
    
    For X = 1 To w配列数
        For j = X + 1 To w配列数
            If Format(G借入金テーブル(X).実際年月日, "yyyy/mm/dd") _
                > Format(G借入金テーブル(j).実際年月日, "yyyy/mm/dd") Then
                
                p借入金テーブル = G借入金テーブル(X)
                
                G借入金テーブル(X) = G借入金テーブル(j)
                         
                G借入金テーブル(j) = p借入金テーブル
            Else
            
                If G借入金テーブル(X).実際年月日 = G借入金テーブル(j).実際年月日 _
                    And G借入金テーブル(X).返済回数 < G借入金テーブル(j).返済回数 Then
                    
                    p借入金テーブル = G借入金テーブル(X)
                    
                    G借入金テーブル(X) = G借入金テーブル(j)
                    
                    G借入金テーブル(j) = p借入金テーブル
                Else
                    If G借入金テーブル(X).実際年月日 = G借入金テーブル(j).実際年月日 _
                        And G借入金テーブル(X).返済回数 = G借入金テーブル(j).返済回数 _
                        And G借入金テーブル(X).据置x回目 > G借入金テーブル(j).据置x回目 Then
                        
                        p借入金テーブル = G借入金テーブル(X)
                        
                        G借入金テーブル(X) = G借入金テーブル(j)
                    
                        G借入金テーブル(j) = p借入金テーブル
                        
                    End If
                    
                End If
           End If
        Next
        
    Next
    
    '*** 同一返済年月日の時　元金額　保証料　手数料　を　合算する
    Call MBD010_同一返済年月日合算(p解約実行日)                 ' 08/12/24 V189
    
    '*** 融資残高再計算
    w配列数 = UBound(G借入金テーブル)                           '08/12/10 V189
    p融資残高 = p借入金マスタ.融資金額
    
    For X = 1 To w配列数
        p借入金テーブル.実際年月日 = G借入金テーブル(X).実際年月日
        p融資残高 = p融資残高 - G借入金テーブル(X).元金額
        G借入金テーブル(X).融資残高 = p融資残高
        
        '*前月融資残高が　０　の時　元金額　利息額　返済金額　融資残高　等　に　０をセット
        If X > 1 And G借入金テーブル(X - 1).融資残高 = 0 Then
            G借入金テーブル(X).元金額 = 0
            G借入金テーブル(X).利息額 = 0
            G借入金テーブル(X).返済金額 = 0
            G借入金テーブル(X).融資残高 = 0
        End If
        
            
        
        '*融資残高マイナスの調整
        If (G借入金テーブル(X).据置x回目 = 2 Or G借入金テーブル(X).据置x回目 = 4) _
            And G借入金テーブル(X).融資残高 < 0 Then
            G借入金テーブル(X).元金額 = G借入金テーブル(X).元金額 + G借入金テーブル(X).融資残高
            w融資残高 = 0
            G借入金テーブル(X).融資残高 = w融資残高
        End If
        
        '*最終返済の時　融資残高＝０　に成るように調整
        If G借入金テーブル(X).実際年月日 = p借入金マスタ.最終返済実行日 _
            And (G借入金テーブル(X).据置x回目 = 2 Or G借入金テーブル(X).据置x回目 = 4) _
            And G借入金テーブル(X).融資残高 <> 0 Then
            G借入金テーブル(X).元金額 = G借入金テーブル(X).元金額 + G借入金テーブル(X).融資残高
            w融資残高 = 0
            G借入金テーブル(X).融資残高 = w融資残高
        End If
        
        '***内入処理の時　利息再計算
        If G借入金テーブル(X).据置x回目 = 1 Or G借入金テーブル(X).据置x回目 = 3 Then
        
            '***内入の利息計算年月日セット及び利息参照
            'G借入金テーブル(X).利息計算年月日 = G借入金テーブル(X).実際年月日   '10/01/30
            '*** 変動金利　設定
            If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then   ' 10/01/23
                If Not IsNull(w解約実行日) And _
                    Format(G借入金テーブル(X).実際年月日, "yyyy/mm/dd") = Format(w解約実行日, "yyyy/mm/dd") Then
                    w利息計算基準年月日 = DateAdd("d", -1, G借入金テーブル(X).利息計算年月日)   '10/01/23
                Else                                                                '10/01/23
                    w利息計算基準年月日 = G借入金テーブル(X).利息計算年月日                     '10/01/23
                End If                                                              '10/01/23
                
                '利息計算基準年月日 = G借入金テーブル(X).利息計算年月日     ' 10/01/23
            Else                                                                ' 10/01/23
                w利息計算基準年月日 = DateAdd("d", -G借入金テーブル(X).利息対象期間日数, G借入金テーブル(X).利息計算年月日) ' 10/01/23
            End If
            
            
            'w利息計算基準年月日 = G借入金テーブル(X).利息計算年月日             '10/01/04
            G借入金テーブル(X).利率 = MBD010_金利参照(p借入金マスタ, w利息計算基準年月日) '10/01/04
            
            G借入金テーブル(X).利息額 = MBD010_利息計算小数点5桁(G借入金テーブル(X).利率, _
                                                               G借入金テーブル(X).元金額, _
                              G借入金テーブル(X).日割日数, p借入金マスタ.金利計算年間日数) '10/01/04
            
            
            'If p借入金マスタ.金利計算年間日数 = 0 Then
            '    G借入金テーブル(X).利息額 = Fix(G借入金テーブル(X).元金額 * _
            '           CCur(G借入金テーブル(X).利率) * G借入金テーブル(X).日割日数 / 36500)
            'Else
            '    G借入金テーブル(X).利息額 = Fix(G借入金テーブル(X).元金額 * _
            '           CCur(G借入金テーブル(X).利率) * G借入金テーブル(X).日割日数 / 36000)
            'End If
        End If
        
        
        '***内入以外の時　利息再計算
        If G借入金テーブル(X).据置x回目 = 2 Or G借入金テーブル(X).据置x回目 = 4 Then
            If p借入金マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
                '利息先払
                G借入金テーブル(X).利息額 = MBD010_利息計算小数点5桁(G借入金テーブル(X).利率, _
                                                               G借入金テーブル(X).融資残高, _
                              G借入金テーブル(X).日割日数, p借入金マスタ.金利計算年間日数) '10/01/04
            
                'If p借入金マスタ.金利計算年間日数 = 0 Then
                '    G借入金テーブル(X).利息額 = Fix(G借入金テーブル(X).融資残高 * _
                '        CCur(G借入金テーブル(X).利率) * G借入金テーブル(X).日割日数 / 36500)
                'Else
                '    G借入金テーブル(X).利息額 = Fix(G借入金テーブル(X).融資残高 * _
                '        CCur(G借入金テーブル(X).利率) * G借入金テーブル(X).日割日数 / 36000)
                'End If
                
            Else
                '利息後払
                G借入金テーブル(X).利息額 = MBD010_利息計算小数点5桁(G借入金テーブル(X).利率, _
                                     G借入金テーブル(X).融資残高 + G借入金テーブル(X).元金額, _
                              G借入金テーブル(X).日割日数, p借入金マスタ.金利計算年間日数) '10/01/04
            
                'If p借入金マスタ.金利計算年間日数 = 0 Then
                '    G借入金テーブル(X).利息額 = Fix((G借入金テーブル(X).融資残高 + G借入金テーブル(X).元金額) _
                '        * CCur(G借入金テーブル(X).利率) * G借入金テーブル(X).日割日数 / 36500)
                'Else
                '    G借入金テーブル(X).利息額 = Fix((G借入金テーブル(X).融資残高 + G借入金テーブル(X).元金額) _
                '        * CCur(G借入金テーブル(X).利率) * G借入金テーブル(X).日割日数 / 36000)
                'End If
            End If
        End If
        
        G借入金テーブル(X).返済金額 = G借入金テーブル(X).元金額 + G借入金テーブル(X).利息額
        
        
        
    Next
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_融資残高再計算_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_融資残高再計算() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub
'
'------------------------------------------------
' MBD010_借入金入力明細Read
'------------------------------------------------
Public Sub MBD010_借入金入力明細Read(p借入金 As MAA910_借入金)
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    
    Dim wi01 As Integer
    Dim wi02 As Integer                 '10/11/21 V189R
    
    Dim w本日年月日 As Variant          '10/11/11 V189R
    Dim X As Integer                    '10/11/11 V189R
    Dim w新利率 As Double               '10/11/11 V189R
    Dim w新利息額 As Double             '10/11/11 V189R
    
    Dim p金利SM利率 As MAA070_金利SM率
'
    On Error GoTo MBD010_借入金入力明細Read_ERR
'
    wi01 = 0
    ReDim G借入金入力(wi01)
    
    p借入金.支払回数 = 0
'
    wstr = ""
    wstr = wstr & "Select"
    wstr = wstr & " 実際年月日,利息計算年月日,元金額,利率,利息額,仮計上利息額,返済金額,融資残高,日割日数,利息対象期間日数"
    
    If p借入金.借入貸付 = P8.FCDbl(XMXA020_区分("借入貸付", "借入")) Then
        wstr = wstr & " From DBDA010_借入金明細TR"
    Else
        wstr = wstr & " From DBDA010_貸付金明細TR"
    End If
    
    wstr = wstr & " Where 借入番号='" & p借入金.借入番号 & "'"
    wstr = wstr & " And 取消フラグ=0"
    wstr = wstr & " And 取消フラグ２=0"
    wstr = wstr & " Order by 返済予定年月,実際年月日"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    Do Until wRs.EOF
    
        wi01 = wi01 + 1
        ReDim Preserve G借入金入力(wi01)
        
        GVar1 = wRs("実際年月日")
        If IsDate(GVar1) Then
            G借入金入力(wi01).借入返済年月日 = CDate(GVar1)
        End If
        GVar1 = wRs("利息計算年月日")
        If IsDate(GVar1) Then
            G借入金入力(wi01).利息計算年月日 = CDate(GVar1)
        End If
        G借入金入力(wi01).元金 = wRs("元金額")
        G借入金入力(wi01).利率 = wRs("利率")
        G借入金入力(wi01).利息額 = wRs("利息額")
        G借入金入力(wi01).仮計上利息額 = wRs("仮計上利息額")
        G借入金入力(wi01).返済金額 = wRs("返済金額")
        G借入金入力(wi01).融資残高 = wRs("融資残高")
        G借入金入力(wi01).日割日数 = wRs("日割日数")        '10/01/08
        G借入金入力(wi01).利息対象期間日数 = wRs("利息対象期間日数")        '10/01/08
        
        '***利息未来シュミュレーション
        p金利SM利率 = MAA070_金利SM率Read(p借入金.金利グループ区分)
        w本日年月日 = Date                                  '10/11/11 V189R
        If p借入金.金利種別 = 0 And G金利SM = True Then     '10/11/11 V189R
            For X = 1 To 100                                '10/11/11 V189R
                If IsNull(p金利SM利率.利率増減率(X).年月日) Then '10/11/11 V189R
                    Exit For                                '10/11/11 V189R
                Else                                        '10/11/11 V189R
                    If Format(G借入金入力(wi01).利息計算年月日, "yyyy/mm/dd") <= Format(w本日年月日, "yyyy/mm/dd") Then '10/11/11 V189R
                        Exit For                            '10/11/11 V189R
                    Else                                    '10/11/11 V189R
                    
                        '**利息後払　1件前の利息計算年月日で計算
                        If Format(p借入金.実行日, "yyyy/mm/dd") >= Format(p金利SM利率.利率増減率(X).年月日, "yyyy/mm/dd") _
                            Or Format(w本日年月日, "yyyy/mm/dd") >= Format(p金利SM利率.利率増減率(X).年月日, "yyyy/mm/dd") Then  '10/11/21 V189R
                            GoTo Ok1                        '10/11/21 V189R
                        End If                              '10/11/21 V189R
                            
                        
                        If p借入金.利息区分 = "2" And wi01 > 1 Then      '10/11/21 V189R
                            wi02 = wi01 - 1                 '10/11/21 V189R
                        Else                                '10/11/21 V189R
                            wi02 = wi01                     '10/11/21 V189R
                        End If                              '10/11/21 V189R
                        
                    
                        If X = 100 Then                     '10/11/11 V189R
                            If Format(G借入金入力(wi02).利息計算年月日, "yyyy/mm/dd") _
                               >= Format(p金利SM利率.利率増減率(X).年月日, "yyyy/mm/dd") Then '10/11/11 V189R
                               
                               If p借入金.利息区分 = "2" And wi01 = 1 Then      '10/11/22 V189R
                                
                               Else                         '10/11/22 V189R
                               
                                    w新利率 = G借入金入力(wi01).利率 + p金利SM利率.利率増減率(X).増減率 '10/11/11 V189R
                                
                                    w新利息額 = Round(G借入金入力(wi01).利息額 * w新利率 / G借入金入力(wi01).利率) '10/11/11 V189R
                                    G借入金入力(wi01).利息額 = w新利息額                '10/11/11 V189R
                                    G借入金入力(wi01).利率 = w新利率                    '10/11/11 V189R
                                    G借入金入力(wi01).返済金額 = G借入金入力(wi01).元金 + G借入金入力(wi01).利息額  '10/11/21 V189R
                               End If                       '10/11/22 V189R
                               
                               Exit For                    '10/11/11 V189R
                            Else                            '10/11/11 V189R
                                GoTo Ok1                    '10/11/11 V189R
                            End If                          '10/11/11 V189R
                        Else                                '10/11/11 V189R
                        
                        
                            If (Format(G借入金入力(wi02).利息計算年月日, "yyyy/mm/dd") >= Format(p金利SM利率.利率増減率(X).年月日, "yyyy/mm/dd")) _
                            And (Format(G借入金入力(wi02).利息計算年月日, "yyyy/mm/dd") < Format(p金利SM利率.利率増減率(X + 1).年月日, "yyyy/mm/dd")) _
                                Or (Format(G借入金入力(wi02).利息計算年月日, "yyyy/mm/dd") >= Format(p金利SM利率.利率増減率(X).年月日, "yyyy/mm/dd") _
                                   And IsNull(p金利SM利率.利率増減率(X + 1).年月日)) Then '10/11/11 V189R
                            
'                            If Format(G借入金入力(wi02).利息計算年月日, "yyyy/mm/dd") _
'                                >= Format(p金利SM利率.利率増減率(X).年月日, "yyyy/mm/dd") Then  '10/11/24
'                                If IsNull(p金利SM利率.利率増減率(X + 1).年月日) Then  '10/11/24 V189R
'                                Else                                                  '10/11/24 V189R
'                                    If Format(G借入金入力(wi02).利息計算年月日, "yyyy/mm/dd") _
'                                        >= Format(p金利SM利率.利率増減率(X).年月日, "yyyy/mm/dd") Then  '10/11/24 V189R
'                                    Else                                                '10/11/24 V189R
'                                        GoTo Ok1                                        '10/11/24 V189R
'                                    End If                                              '10/11/24 V189R
'                                End If                                                  '10/11/24 V189R
                                
                            
                            
                            
                            
                                   
                                If p借入金.利息区分 = "2" And wi01 = 1 Then      '10/11/22 V189R
                                
                                Else                         '10/11/22 V189R
                                    
                                    w新利率 = G借入金入力(wi01).利率 + p金利SM利率.利率増減率(X).増減率 '10/11/11 V189R
                                
                                    w新利息額 = Round(G借入金入力(wi01).利息額 * w新利率 / G借入金入力(wi01).利率) '10/11/11 V189R
                                    G借入金入力(wi01).利息額 = w新利息額                '10/11/11 V189R
                                    G借入金入力(wi01).利率 = w新利率                    '10/11/11 V189R
                                    G借入金入力(wi01).返済金額 = G借入金入力(wi01).元金 + G借入金入力(wi01).利息額  '10/11/21 V189R
                                End If                          '10/11/22 V189R
                                
                                Exit For                        '10/11/11 V189R
                            Else                                '10/11/11 V189R
                        
                                GoTo Ok1                        '10/11/11 V189R
                            End If                              '10/11/11 V189R
                        End If                                  '10/11/11 V189R
                    End If                                      '10/11/11 V189R
                End If                                          '10/11/11 V189R
                
Ok1:
            Next                                                '10/11/11 V189R
            
        End If                                                    '10/11/12 V189R
Ok2:
                
        p借入金.支払回数 = p借入金.支払回数 + 1
        p借入金.変動最終利率 = G借入金入力(wi01).利率
        
        
        wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_借入金入力明細Read_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入金入力明細Read() でエラー" + vbCrLf + vbCrLf + _
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
' MBD010_利息計算年月日
'------------------------------------------------
Public Function MBD010_利息計算年月日(p対象年月 As Variant, p支払日 As Integer, p営業日区分 As Integer, _
                                      p利息計算日数区分 As Integer) As Variant
                                      
 '
    Dim wd01 As Date
'
    On Error GoTo MBD010_利息計算年月日_ERR
'
    If p利息計算日数区分 = 1 Then
        MBD010_利息計算年月日 = C年月日.GetDate("設定", p対象年月, p支払日)
    Else
        wd01 = MXA030_翌営業年月日計算(CDate(p対象年月), p支払日, p営業日区分)
        MBD010_利息計算年月日 = Format(wd01, "yyyy/mm/dd")
    End If
    
    wd01 = MXA030_翌営業年月日計算(CDate(p対象年月), p支払日, p営業日区分)  'V182 2008/01/29
    
    GDate利息対象年月日 = Format(wd01, "yyyy/mm/dd")                        'V182 2008/01/29
    'GDate利息対象年月日 = MBD010_利息計算年月日                                 '09/12/26
    
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_利息計算年月日_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_利息計算年月日() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

'------------------------------------------------
' MBD010_借入手入力残高
'------------------------------------------------
Public Function MBD010_借入金手入力残高(p借入計画マスタ As MAA910_借入金, _
                                            p前当残F As Integer, p残高年月 As Date) As Double
'
    Dim j As Integer
    Dim w残高 As Double
    Dim w対象年月 As Date
    Dim w開始年月 As Date
    Dim w終了年月 As Date
'
    On Error GoTo MBD010_借入金手入力残高_ERR
'
    MBD010_借入金手入力残高 = 0                         ' 07/02/18 V180
    w残高 = 0
    w開始年月 = MBA010_対象年月(CDate(p借入計画マスタ.実行日))
    w終了年月 = MBA010_対象年月(CDate(p借入計画マスタ.最終返済実行日))
    'w終了年月 = p借入計画マスタ.最終返済年月            ' 07/02/18 V180
    
    If p残高年月 > w開始年月 And p残高年月 <= w終了年月 Then
        For j = 1 To UBound(G借入金入力)
            w対象年月 = MBA010_対象年月((CDate(G借入金入力(j).借入返済年月日)))
            If p前当残F = 0 Then
            'Select Case p前当残F
            '    Case 0:
                    If p残高年月 = w対象年月 Then
                        w残高 = G借入金入力(j).融資残高 + G借入金入力(j).元金
                        Exit For
                    Else
                        If p残高年月 < w対象年月 Then
                            w残高 = G借入金入力(j).融資残高 + G借入金入力(j).元金
                            Exit For
                        End If
                        
                    End If
            Else
            
            
            '    Case 1:
                    If p残高年月 < w対象年月 Then
                        w残高 = G借入金入力(j).融資残高 + G借入金入力(j).元金
                        Exit For
                    End If
            'End Select
            End If
            
        Next j
        
        MBD010_借入金手入力残高 = w残高
    End If
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_借入金手入力残高_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入金手入力残高() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

'------------------------------------------------
' MBD010_借入標準入力残高
'------------------------------------------------
Public Function MBD010_借入金標準入力残高(p借入計画マスタ As MAA910_借入金, p金融リストラ As String, _
                                p前当残F As Integer, p残高年月 As Date, p借入金管理区分 As String) As Double
'
    Dim j As Integer
    Dim w残高 As Double
    Dim w対象年月 As Date
    Dim w開始年月 As Date
    Dim w終了年月 As Date
    Dim w解約実行日 As Variant      ' 07/02/26 V180
'
    On Error GoTo MBD010_借入金標準入力残高_ERR
'
    MBD010_借入金標準入力残高 = 0                                             ' 07/02/18 V180
    w残高 = 0
    w開始年月 = MBA010_対象年月(CDate(p借入計画マスタ.実行日))
    w終了年月 = MBA010_対象年月(CDate(p借入計画マスタ.最終返済実行日))
    
    If p残高年月 > w開始年月 And p残高年月 <= w終了年月 Then
        'For J = 1 To UBound(G借入金入力)
        For j = 1 To UBound(G借入金テーブル)
            
            Call MBA010_借入金年月算出(G借入金テーブル(j).返済予定年月, _
                    G借入金テーブル(j).実際年月日, p借入計画マスタ.支払日)  ' 07/02/12 V180
                    
                   
                    
            If p借入金管理区分 = XMXA020_区分("借入金管理区分", "管理用") Then '07/02/18 V180
                w対象年月 = MBA010_対象年月((CDate(G管理年月)))             ' 07/02/18 V180
            Else                                                            ' 07/02/180V180
                w対象年月 = MBA010_対象年月((CDate(G実績年月)))             ' 07/02/18 V180
            End If                                                          ' 07/02/180V180
            
            '解約算出
            If p金融リストラ > "" _
                And p金融リストラ = p借入計画マスタ.金融リストラ番号 Then  ' 07/02/26 V180
                w解約実行日 = p借入計画マスタ.金融解約実行日                ' 07/02/26 V180
            Else                                                            ' 07/02/26 V180
                w解約実行日 = p借入計画マスタ.解約実行日                    ' 07/02/26 V180
            End If                                                          ' 07/02/26 V180
                
            If Format(w解約実行日, "yyyymmdd") = _
                Format(G借入金テーブル(j).実際年月日, "yyyymmdd") Then      ' 07/02/26 V180
                w対象年月 = MBA010_対象年月((CDate(G実績年月)))             ' 07/02/26 V180
            End If                                                          ' 07/02/26 V180
               
            
            
            If p前当残F = 0 Then
            'Select Case p前当残F
            '    Case 0:
                    If p残高年月 = w対象年月 Then
                        w残高 = G借入金テーブル(j).融資残高 + G借入金テーブル(j).元金額 ' 07/02/18 V180
                        Exit For
                    Else
                        If p残高年月 < w対象年月 Then
                            w残高 = G借入金テーブル(j).融資残高 + G借入金テーブル(j).元金額 ' 07/02/18 V180
                            Exit For
                        End If
                        
                    End If
            Else
            
            
            '    Case 1:
                    If p残高年月 < w対象年月 Then
                        w残高 = G借入金テーブル(j).融資残高 + G借入金テーブル(j).元金額 ' 07/02/18 V180
                        Exit For
                    End If
            'End Select
            End If
            
        Next j
        
        MBD010_借入金標準入力残高 = w残高
    End If
'
    Exit Function
'
'----------< ERROR ROUTINE >--------------------------------------------------
MBD010_借入金標準入力残高_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入金標準入力残高() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

'------------------------------------------------
' MBD010_借入最終融資残高
'------------------------------------------------
Public Function MBD010_借入最終融資残高(p借入計画マスタ As MAA910_借入金, p残高年月 As Date) As Double
'
    Dim j As Integer
'
    On Error GoTo MBD010_借入最終融資残高_ERR
'
    MBD010_借入最終融資残高 = 0                                             ' 07/02/18 V180
    
    For j = UBound(G借入金テーブル) To 1 Step -1
        If CDate(G借入金テーブル(j).実際年月日) = CDate(p残高年月) Then
            MBD010_借入最終融資残高 = G借入金テーブル(j).融資残高
            Exit For
        End If
    Next j
'
    Exit Function
'
'----------< ERROR ROUTINE >--------------------------------------------------
MBD010_借入最終融資残高_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入最終融資残高() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

'------------------------------------------------
' MBD010_借入融資残高
'------------------------------------------------
Public Function MBD010_借入指定年月融資残高(p借入計画マスタ As MAA910_借入金, p残高年月 As Date) As Double
'
    Dim j As Integer
'
    On Error GoTo MBD010_借入指定年月融資残高_ERR
'
    MBD010_借入指定年月融資残高 = 0                                             ' 07/02/18 V180
    
    For j = 1 To UBound(G借入金テーブル)
        If CDate(G借入金テーブル(j).実際年月日) <= CDate(p残高年月) Then
            MBD010_借入指定年月融資残高 = G借入金テーブル(j).融資残高
        ElseIf CDate(G借入金テーブル(j).実際年月日) > CDate(p残高年月) Then
            Exit For
        End If
    Next j
'
    Exit Function
'
'----------< ERROR ROUTINE >--------------------------------------------------
MBD010_借入指定年月融資残高_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入指定年月融資残高() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

'------------------------------------------------
' MBD010_借入前月残高
'------------------------------------------------
Public Function MBD010_借入前月残高(p借入金 As MAA910_借入金, p年月日 As Date) As Double
'
    Dim j As Integer
    Dim wd開始年月 As Date, wd終了年月 As Date
'
    On Error GoTo MBD010_借入前月残高_ERR
'
    MBD010_借入前月残高 = 0
    
    wd開始年月 = CDate(p借入金.実行日)
    wd終了年月 = CDate(p借入金.最終返済実行日)
    
    If p年月日 > wd開始年月 And p年月日 <= wd終了年月 Then
        For j = UBound(G借入金テーブル) To 1 Step -1
            If p年月日 > G借入金テーブル(j).実際年月日 Then
                MBD010_借入前月残高 = G借入金テーブル(j).融資残高
                Exit For
            End If
        Next j
    End If
'
    Exit Function
'
'----------< ERROR ROUTINE >--------------------------------------------------
MBD010_借入前月残高_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入前月残高() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

'---------------------------------------------------------------------
'　　MBD010_借入金据置期間一括利息支払  08/03/14 V185
'---------------------------------------------------------------------
Public Function MBD010_借入金据置期間一括利息支払(p返済予定年月 As Variant, _
                                                  p一括支払利息開始年月 As Date, _
                                                  p初回返済年月 As Variant, _
                                                  p利息区分 As String, _
                                                  p返済単位月数 As Integer) As Integer
'
    Dim w一括支払利息開始年月 As Date
'
    On Error GoTo MBD010_借入金据置期間一括利息支払_ERR
'
    w一括支払利息開始年月 = p一括支払利息開始年月
    
    MBD010_借入金据置期間一括利息支払 = 0
    
    
    
    Select Case p利息区分
        Case "1"
        
Loop1:
            If Format(w一括支払利息開始年月, "yyyymmdd") >= Format(p初回返済年月, "yyyymmdd") Then
                GoTo Ok1
            End If
            
            If Format(p返済予定年月, "yyyymmdd") = Format(w一括支払利息開始年月, "yyyymmdd") Then
                MBD010_借入金据置期間一括利息支払 = 1
                GoTo Ok1
            End If
            
            w一括支払利息開始年月 = DateAdd("m", p返済単位月数, w一括支払利息開始年月)
            GoTo Loop1
Ok1:
            
            
        Case "2"
        
Loop2:
            If Format(w一括支払利息開始年月, "yyyymmdd") > Format(p初回返済年月, "yyyymmdd") Then
                GoTo Ok2
            End If
            
            If Format(p返済予定年月, "yyyymmdd") = Format(w一括支払利息開始年月, "yyyymmdd") Then
                MBD010_借入金据置期間一括利息支払 = 1
                GoTo Ok2
            End If
            
            w一括支払利息開始年月 = DateAdd("m", p返済単位月数, w一括支払利息開始年月)
            GoTo Loop2
            
Ok2:
            
    End Select
'
    Exit Function
'
'----------< ERROR ROUTINE >--------------------------------------------------
MBD010_借入金据置期間一括利息支払_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入金据置期間一括利息支払() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

'---------------------------------------------------------------------
'　　MBD010_借入金解約据置期間一括利息支払日  08/03/14 V185
'---------------------------------------------------------------------
Public Function MBD010_借入金解約据置期間一括利息支払日(p解約実行日 As Variant, _
                                                  p実行日 As Variant, _
                                                  p返済予定年月 As Date, _
                                                  p一括支払利息開始年月 As Date, _
                                                  p初回返済年月 As Variant, _
                                                  p利息区分 As String, _
                                                  p返済単位月数 As Integer, _
                                                  p支払日 As Integer, _
                                                  p営業日区分 As Integer, _
                                                  p利息計算日数区分 As Integer) As Date     '10/02/02
'
    Dim w一括支払利息開始年月 As Date
    Dim p初回返済年月日 As Variant
    Dim w一括支払利息開始年月日 As Date
    Dim wx一括支払利息開始年月日 As Date
    Dim w年月日 As Date                 '10/02/02
    Dim sv解約実行日 As Date            '10/02/02
    
    Dim w初回利息支払年月日 As Date     '08/03/16 V185
'
    On Error GoTo MBD010_借入金解約据置期間一括利息支払日_ERR
'
    '***解約実行日を利息計算日に換算する    10/02/02
    sv解約実行日 = p解約実行日              '10/02/02
    w年月日 = MXA030_翌営業年月日計算(p返済予定年月, p支払日, p営業日区分)  '10/02/02
    If Format(w年月日, "yyyy/mm/dd") = Format(p解約実行日, "yyyy/mm/dd") Then '10/02/02
        p解約実行日 = MBD010_利息計算年月日(p返済予定年月, p支払日, _
                                                            p営業日区分, p利息計算日数区分) '10/02/02
    End If                                                                  '10/02/02
    
'
    w初回利息支払年月日 = MBD010_利息計算年月日(p一括支払利息開始年月, p支払日, _
                                                            p営業日区分, p利息計算日数区分)
                                                            
     
    w一括支払利息開始年月 = p一括支払利息開始年月
    
    
    Select Case p利息区分
        Case "1"
Loop1:
            w一括支払利息開始年月日 = MBD010_利息計算年月日(w一括支払利息開始年月, p支払日, _
                                                            p営業日区分, p利息計算日数区分)
                                                            
            wx一括支払利息開始年月日 = DateAdd("m", p返済単位月数, w一括支払利息開始年月)
            wx一括支払利息開始年月日 = MBD010_利息計算年月日(wx一括支払利息開始年月日, p支払日, _
                                                            p営業日区分, p利息計算日数区分)
                                                            
            If Format(p解約実行日, "yyyymmdd") >= Format(p実行日, "yyyymmdd") _
               And Format(p解約実行日, "yyyymmdd") <= Format(w一括支払利息開始年月日, "yyyymmdd") Then
               
                MBD010_借入金解約据置期間一括利息支払日 = w一括支払利息開始年月日
                
                GoTo Ok1
            End If
                                                            
                                                            
                                                            
            If Format(p解約実行日, "yyyymmdd") >= Format(w一括支払利息開始年月日, "yyyymmdd") _
               And Format(p解約実行日, "yyyymmdd") <= Format(wx一括支払利息開始年月日, "yyyymmdd") Then
               
                MBD010_借入金解約据置期間一括利息支払日 = wx一括支払利息開始年月日
                
                GoTo Ok1
            End If
               
               
            w一括支払利息開始年月 = DateAdd("m", p返済単位月数, w一括支払利息開始年月)
                
            GoTo Loop1
            
Ok1:
            
        Case "2"
Loop2:

            w一括支払利息開始年月日 = MBD010_利息計算年月日(w一括支払利息開始年月, p支払日, _
                                                            p営業日区分, p利息計算日数区分)
                                                            
            wx一括支払利息開始年月日 = DateAdd("m", p返済単位月数, w一括支払利息開始年月)
            wx一括支払利息開始年月日 = MBD010_利息計算年月日(wx一括支払利息開始年月日, p支払日, _
                                                            p営業日区分, p利息計算日数区分)
                                                            
            If Format(p解約実行日, "yyyymmdd") >= Format(p実行日, "yyyymmdd") _
               And Format(p解約実行日, "yyyymmdd") <= Format(w初回利息支払年月日, "yyyymmdd") Then
               
                MBD010_借入金解約据置期間一括利息支払日 = p実行日   '08/03/16 V185
                
                GoTo Ok2
            End If
                                                            
                                                            
             
            If Format(p解約実行日, "yyyymmdd") >= Format(w一括支払利息開始年月日, "yyyymmdd") _
               And Format(p解約実行日, "yyyymmdd") <= Format(wx一括支払利息開始年月日, "yyyymmdd") Then
                    
                MBD010_借入金解約据置期間一括利息支払日 = w一括支払利息開始年月日
                
                GoTo Ok2
            End If
            
            w一括支払利息開始年月 = DateAdd("m", p返済単位月数, w一括支払利息開始年月)
                
            GoTo Loop2
            
Ok2:


    End Select
    
    p解約実行日 = sv解約実行日              '10/02/02
'
    Exit Function
'
'----------< ERROR ROUTINE >--------------------------------------------------
MBD010_借入金解約据置期間一括利息支払日_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入金解約据置期間一括利息支払日() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

'------------------------------------------------
' MBD010_実行解約年月
'------------------------------------------------
Public Function MBD010_実行解約年月(p年月日 As Variant, _
                                    p初回返済年月 As Variant, p初回返済実行日 As Variant, _
                                    p最終返済年月 As Variant, p最終返済実行日 As Variant, _
                                    p支払日 As Integer, p営業日区分 As Integer) As Variant ' 07/01/30 V180
'
    Dim wdd As Integer
    Dim w年月(4) As Date
    Dim w年月日(4) As Date
    Dim j As Integer
'
    On Error GoTo MBD010_実行解約年月_ERR
'
    '***解約年月＆解約年月日　テーブル
    If IsNull(p年月日) Then
        MBD010_実行解約年月 = Null
    Else
    '***解約年月＆解約年月日　テーブル
        w年月(4) = C年月日.GetDate("月始", CDate(p年月日))
        w年月(4) = DateAdd("m", 1, w年月(4)) ' 08/07/27 V188
        If Day(p年月日) > p支払日 Then
            w年月(4) = DateAdd("m", 1, w年月(4))
        End If
        
        For j = 4 To 1 Step -1
            If Format(w年月(j), "yyyy/mm/dd") = Format(p初回返済年月, "yyyy/mm/dd") Then
                w年月日(j) = p初回返済実行日
                GoTo next1
            End If
            If Format(w年月(j), "yyyy/mm/dd") = Format(p最終返済年月, "yyyy/mm/dd") Then
                w年月日(j) = p最終返済実行日
                GoTo next1
            End If
            
            w年月日(j) = MXA030_翌営業年月日計算(w年月(j), p支払日, p営業日区分)
next1:
            w年月(j - 1) = DateAdd("m", -1, w年月(j))
        Next
        
        '*** 解約年月算出
        For j = 1 To 3
            If Format(p年月日, "yyyy/mm/dd") > Format(w年月日(j), "yyyy/mm/dd") _
               And Format(p年月日, "yyyy/mm/dd") <= Format(w年月日(j + 1), "yyyy/mm/dd") Then
               MBD010_実行解約年月 = w年月(j + 1)
               Exit For
            End If
        Next
        
    End If
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_実行解約年月_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_実行解約年月() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

'------------------------------------------------
' MBD010_利息計算小数点5桁
'------------------------------------------------
Public Function MBD010_利息計算小数点5桁(p利率 As Double, p金額 As Double, p日数 As Integer, _
                                                p利息計算日数区分 As Integer) As Double
'
    Dim w利率A As Double
'
    On Error GoTo MBD010_利息計算小数点5桁_ERR
    
    w利率A = Fix((p利率 + 0.000001) * 100000)
    If p利息計算日数区分 = 0 Then
        MBD010_利息計算小数点5桁 = Fix((p金額 / 365000) * w利率A * p日数 / 10000) '10/06/19 V195
        
    Else
        MBD010_利息計算小数点5桁 = Fix((p金額 / 360000) * w利率A * p日数 / 10000) '10/06+/19 V195
        
    End If
'
''
'    Dim w利率 As Double
'    Dim w利率A As Double
'    Dim w利率B As Double
'
''
'    On Error GoTo MBD010_利息計算小数点5桁_ERR
'
'    w利率A = Fix((p利率 + 0.000001) * 100000)
'    w利率B = Fix(w利率A / 10)
'    w利率B = Fix(w利率B * 10)
'    If w利率A <> w利率B Then
'        w利率A = w利率A / 10000
'        If p利息計算日数区分 = 0 Then
'            MBD010_利息計算小数点5桁 = Fix(p金額 * CCur(w利率A) * p日数 / 365000)
'
'        Else
'            MBD010_利息計算小数点5桁 = Fix(p金額 * CCur(w利率A) * p日数 / 360000)
'
'        End If
'
'    Else
'
'        If p利息計算日数区分 = 0 Then
'            MBD010_利息計算小数点5桁 = Fix(p金額 * CCur(p利率) * p日数 / 36500)
'
'        Else
'            MBD010_利息計算小数点5桁 = Fix(p金額 * CCur(p利率) * p日数 / 36000)
'
'        End If
'    End If
''
    
'
    Exit Function
'
'----------< ERROR ROUTINE >--------------------------------------------------
MBD010_利息計算小数点5桁_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_利息計算小数点5桁() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

'------------------------------------------------
' MBD010_実行日支払年月算出
'------------------------------------------------
Public Function MBD010_実行日支払年月算出(p実行日 As Variant, p営業日区分 As Integer, p支払日 As Integer) As Variant
'
    Dim w支払日 As Integer
    Dim wyy As Integer
    Dim wmm As Integer
    Dim wv01 As Variant
'
    On Error GoTo MBD010_実行日支払年月算出_ERR
    
    w支払日 = Day(C年月日.GetDate("月末", p実行日))
    
    '*** 実行日の締年月算出 10/01/16
    wv01 = Format(p実行日, "yyyy/mm/01")  '10/01/16
    If p支払日 <= w支払日 Then
        w支払日 = p支払日
    End If
    
    If p営業日区分 = 0 Then                           '10/01/20
        wyy = Year(p実行日)                            '10/01/20
        wmm = Month(p実行日)                           '10/01/20
        wv01 = Right("0000" & CStr(wyy), 4) & "/" & Right("00" & CStr(wmm), 2) & "/" & Right("00" & CStr(w支払日), 2)
        wv01 = Format(wv01, "yyyy/mm/dd")         '10/01/20
    Else                                                            '10/01/20
        wv01 = MXA030_翌営業年月日計算(CDate(wv01), p支払日, p営業日区分)  '10/01/20
    End If                                                          '10/01/20
    
    
    If Format(wv01, "yyyy/mm/dd") <= Format(p実行日, "yyyy/mm/dd") Then '10/01/16
        wv01 = DateAdd("m", 1, CDate(wv01))                                          '10/01/16
    End If                  '10/01/16
    MBD010_実行日支払年月算出 = Format(wv01, "yyyy/mm/01")     '10/01/16
'
    Exit Function
'
'----------< ERROR ROUTINE >--------------------------------------------------
MBD010_実行日支払年月算出_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_実行日支払年月算出() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

'------------------------------------------------
' MBD010_同一返済年月日合算
'------------------------------------------------
Private Sub MBD010_同一返済年月日合算(p解約実行日 As Variant)
'
    Dim p借入金テーブル() As MAA910_借入金テーブル
    Dim sv借入金テーブル As MAA910_借入金テーブル
    Dim wFind As Boolean
    Dim w配列数 As Integer
    Dim j As Integer
    Dim Cnt As Integer
    Dim wsw As Integer
    Dim w返済回数 As Integer                ' 08/12/16 V189
    Dim w日割日数 As Integer                '10/01/30
    Dim w利息対象期間日数 As Integer        '10/01/30
    Dim w利息計算年月日 As Date             '10/01/30
'
    On Error GoTo MBD010_同一返済年月日合算_ERR
'
    p借入金テーブル = G借入金テーブル
    
    ReDim G借入金テーブル(0)
    
    w配列数 = UBound(p借入金テーブル)
    Cnt = 0
    wsw = 0
    
    For j = 1 To w配列数
        
        If wsw = 0 Then
            wsw = 1
SET1:
            sv借入金テーブル = p借入金テーブル(j)
            GoTo next1
            
        Else
            If Not IsNull(p解約実行日) Then                                     ' 08/12/24 V189
                If Format(p借入金テーブル(j).実際年月日, "yyyy/mm/dd") _
                    > Format(p解約実行日, "yyyy/mm/dd") Then                    ' 08/12/24 V189
                    GoTo next1                                                  ' 08/12/24 V189
                End If                                                          ' 08/12/24 V189
            End If                                                              ' 08/12/24 V189
            
            If (Format(p借入金テーブル(j).実際年月日, "yyyy/mm/dd") _
                     = Format(sv借入金テーブル.実際年月日, "yyyy/mm/dd")) Then   ' 08/12/24 V189
                     
                If p借入金テーブル(j).据置x回目 = 2 Then               '10/01/30
                    sv借入金テーブル.利息計算年月日 = p借入金テーブル(j).利息計算年月日 '10/01/30
                End If                                                          '10/01/30
                
                If p借入金テーブル(j).据置x回目 = 2 And p借入金テーブル(j).日割日数 <> 0 Then '10/01/30
                    sv借入金テーブル.据置x回目 = 4                                      '10/01/30
                    sv借入金テーブル.日割日数 = p借入金テーブル(j).日割日数             '10/01/30
                    sv借入金テーブル.利息対象期間日数 = p借入金テーブル(j).利息対象期間日数 '10/01/30
                End If                                                                  '10/01/30
                
                
                
                
                If p借入金テーブル(j).据置x回目 = 1 Or p借入金テーブル(j).据置x回目 = 3 Then
                    If p借入金テーブル(j - 1).日割日数 = 0 Then ' 08/12/14 V189
                        w返済回数 = sv借入金テーブル.返済回数   ' 08/12/16 V189
                        'w日割日数 = sv借入金テーブル.日割日数   '10/01/30
                        'w利息対象期間日数 = sv借入金テーブル.利息対象期間日数 '10/01/30
                        w利息計算年月日 = sv借入金テーブル.利息計算年月日       '10/01/30
                        sv借入金テーブル = p借入金テーブル(j)   ' 08/12/14 V189
                        sv借入金テーブル.返済回数 = w返済回数   ' 08/12/16 V189
                        'sv借入金テーブル.日割日数 = w日割日数   '10/01/30
                        'sv借入金テーブル.利息対象期間日数 = w利息対象期間日数   '10/01/30
                        sv借入金テーブル.利息計算年月日 = w利息計算年月日       '10/01/30
                        
                        GoTo next1                              ' 08/12/14 V189
                    Else                                        ' 08/12/04 V189
                        sv借入金テーブル.据置x回目 = 4
                        sv借入金テーブル.元金額 = p借入金テーブル(j).元金額
                    End If                                      ' 08/12/14 V189
                End If
                
                sv借入金テーブル.保証料 = sv借入金テーブル.保証料 + p借入金テーブル(j).保証料
                sv借入金テーブル.手数料 = sv借入金テーブル.手数料 + p借入金テーブル(j).手数料
                
                '*** 解約の時　ウチイレ元金額を　ゼロにする　調整
                If Format(p解約実行日, "yyyy/mm/dd") = _
                    Format(p借入金テーブル(j).実際年月日, "yyyy/mm/dd") Then    '10/02/01
                    sv借入金テーブル.元金額 = 0                                 '10/02/01
                End If                                                          '10/02/01
                
                GoTo next1
            Else
                Cnt = Cnt + 1
                ReDim Preserve G借入金テーブル(Cnt)
                G借入金テーブル(Cnt) = sv借入金テーブル
                GoTo SET1
            End If
        End If
        
next1:
    Next
    
    Cnt = Cnt + 1
    ReDim Preserve G借入金テーブル(Cnt)
    G借入金テーブル(Cnt) = sv借入金テーブル
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_同一返済年月日合算_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_同一返済年月日合算() でエラー" + vbCrLf + vbCrLf + _
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
' MBD010_借入金入力明細作成_日割日数再計算
'------------------------------------------------
Public Sub MBD010_借入金入力明細作成_日割日数再計算(p借入金 As MAA910_借入金, pTbl As String)
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    Dim wi01 As Integer
    Dim wDate1 As Date, wDate2 As Date, wDate3 As Date
'
    wstr = "Update " & pTbl
    wstr = wstr & " Set 日割日数 = 0"
    wstr = wstr & " Where 借入番号 = '" & p借入金.借入番号 & "'"
    GDb.Execute wstr
'
    If p借入金.利息区分 = XMXA020_区分("利息区分", "利息先払") And Not IsNull(P8.FCDate(p借入金.最終返済実行日)) Then
        wDate2 = Format(p借入金.最終返済実行日, "yyyy/mm/dd")
    
        wstr = "Select * From " & pTbl
        wstr = wstr & " Where 借入番号 = '" & p借入金.借入番号 & "'"
        wstr = wstr & " And 元金額+利息額<>0"
        wstr = wstr & " And Format(実際年月日, 'yyyy/mm/dd') < '" & Format(p借入金.最終返済実行日, "yyyy/mm/dd") & "'"
        wstr = wstr & " order by 利息計算年月日 desc"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.EOF
            wDate3 = Format(wRs("実際年月日"), "yyyy/mm/dd")
            
            wDate1 = Format(wRs("利息計算年月日"), "yyyy/mm/dd")
            wi01 = DateDiff("d", wDate1, wDate2)
            
            If Format(wDate3, "yyyy/mm/dd") = Format(p借入金.実行日, "yyyy/mm/dd") Then
            'wdate1が実行日の場合
                
                '実行日を含めた日数
                wi01 = wi01 + 1
                
                '実行日控除
                If p借入金.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日控除")) _
                Or p借入金.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) Then
                    wi01 = wi01 - 1
                End If
            End If
            
            If Format(wDate2, "yyyy/mm/dd") = Format(p借入金.最終返済実行日, "yyyy/mm/dd") Then
            '最終返済日より前の1件目
                
                '最終返済実行日を除く日数
                'wi01 = wi01 - 1
                
                '最終返済日控除
                If p借入金.利息控除区分 = CDbl(XMXA020_区分("利息控除", "最終返済日控除")) _
                Or p借入金.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) Then
                    wi01 = wi01 - 1
                End If
            End If
            
            wRs("日割日数") = wi01
            
            wDate2 = wDate1
            
            wRs.Update
        
            wRs.MoveNext
        Loop
        wRs.Close
        Set wRs = Nothing
        
    Else
    
        wDate1 = Format(p借入金.実行日, "yyyy/mm/dd")
    
        wstr = "Select * From " & pTbl
        wstr = wstr & " Where 借入番号 = '" & p借入金.借入番号 & "'"
        wstr = wstr & " And 元金額+利息額<>0"
        wstr = wstr & " And Format(実際年月日, 'yyyy/mm/dd') > '" & Format(p借入金.実行日, "yyyy/mm/dd") & "'"
        wstr = wstr & " order by 利息計算年月日"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.EOF
            wDate3 = Format(wRs("実際年月日"), "yyyy/mm/dd")
            
            wDate2 = Format(wRs("利息計算年月日"), "yyyy/mm/dd")
            wi01 = DateDiff("d", wDate1, wDate2)
            
            If Format(wDate1, "yyyy/mm/dd") = Format(p借入金.実行日, "yyyy/mm/dd") Then
            '実行日移行の1件目
                
                '実行日を含めた日数
                wi01 = wi01 + 1
                
                '実行日控除
                If p借入金.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日控除")) _
                Or p借入金.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) Then
                    wi01 = wi01 - 1
                End If
            End If
            
            If Format(wDate3, "yyyy/mm/dd") = Format(p借入金.最終返済実行日, "yyyy/mm/dd") Then
            'wdate1が最終返済日の場合
                
                '最終返済日控除
                If p借入金.利息控除区分 = CDbl(XMXA020_区分("利息控除", "最終返済日控除")) _
                Or p借入金.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) Then
                    wi01 = wi01 - 1
                End If
            End If
            
            wRs("日割日数") = wi01
            
            wDate1 = wDate2
            
            wRs.Update
        
            wRs.MoveNext
        Loop
        wRs.Close
        Set wRs = Nothing
    End If
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_借入金入力明細作成_日割日数再計算_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入金入力明細作成_日割日数再計算() でエラー" + vbCrLf + vbCrLf + _
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
' MBD010_借入金入力明細作成_利率変更
'------------------------------------------------
Public Sub MBD010_借入金入力明細作成_利率変更(p借入データ As MAA910_借入金, pTbl As String, pDate As Date, _
                                                p利率 As Double, p利率New As Double)
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
'
    wstr = "Select * From " & pTbl
    wstr = wstr & " Where 借入番号 = '" & p借入データ.借入番号 & "'"
    wstr = wstr & " And Format(実際年月日, 'yyyy/mm/dd') > '" & Format(pDate, "yyyy/mm/dd") & "'"
    wstr = wstr & " order by 実際年月日"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    Do Until wRs.EOF
        
        If p利率 = P8.FCDbl(wRs("利率")) Then
            wRs("利率") = p利率New
        Else
            Exit Do
        End If
        
        wRs.Update
    
        wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_借入金入力明細作成_利率変更_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入金入力明細作成_利率変更() でエラー" + vbCrLf + vbCrLf + _
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
' MBD010_借入金入力明細作成_利息額再計算
'------------------------------------------------
Public Sub MBD010_借入金入力明細作成_利息額再計算(p借入データ As MAA910_借入金, pTbl As String, pDate As Date)
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
'
    ReDim G借入金入力(0)
    
    wstr = "Select * From " & pTbl
    wstr = wstr & " Where 借入番号 = '" & p借入データ.借入番号 & "'"
    wstr = wstr & " And Format(実際年月日, 'yyyy/mm/dd') >= '" & Format(pDate, "yyyy/mm/dd") & "'"
    wstr = wstr & " order by 実際年月日"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    Do Until wRs.EOF
        
        G借入金入力(0).元金 = wRs("元金額")
        G借入金入力(0).利率 = wRs("利率")
        G借入金入力(0).融資残高 = wRs("融資残高")
        G借入金入力(0).日割日数 = wRs("日割日数")
        
        '利息額自動計算
        G借入金入力(0).利息額 = 0
        If p借入データ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
            G借入金入力(0).利息額 = MBD010_利息計算小数点5桁(G借入金入力(0).利率, _
                        G借入金入力(0).融資残高, G借入金入力(0).日割日数, p借入データ.金利計算年間日数)
        Else
            G借入金入力(0).利息額 = MBD010_利息計算小数点5桁(G借入金入力(0).利率, _
                        G借入金入力(0).元金 + G借入金入力(0).融資残高, G借入金入力(0).日割日数, p借入データ.金利計算年間日数)
        End If
        
        wRs("利息額") = G借入金入力(0).利息額
        wRs("返済金額") = G借入金入力(0).元金 + G借入金入力(0).利息額
        
        
        wRs.Update
    
        wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_借入金入力明細作成_利息額再計算_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入金入力明細作成_利息額再計算() でエラー" + vbCrLf + vbCrLf + _
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
' MBD010_借入金入力明細_利息額自動計算
'------------------------------------------------
Public Sub MBD010_借入金入力明細_利息額自動計算(p借入データ As MAA910_借入金, pTbl As String)
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    
    Dim j As Integer, wi01 As Integer
'
    On Error GoTo MBD010_借入金入力明細_利息額自動計算_ERR
'
    wi01 = 0
    ReDim G借入金入力(wi01)
'
    '** 明細ファイル 削除 **
    wstr = ""
    wstr = wstr + "Delete * From DCDA020_借入金明細"
    GDb.Execute wstr
'
    wstr = ""
    wstr = wstr & "Select"
    wstr = wstr & " 実際年月日,利息計算年月日,元金額,利率,利息額,仮計上利息額,返済金額,融資残高,日割日数,利息対象期間日数"
    wstr = wstr & " From " & pTbl
    wstr = wstr & " Where 借入番号='" & p借入データ.借入番号 & "'"
    wstr = wstr & " And 取消フラグ=0"
    wstr = wstr & " And 取消フラグ２=0"
    wstr = wstr & " Order by 返済予定年月,実際年月日"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    Do Until wRs.EOF
    
        wi01 = wi01 + 1
        ReDim Preserve G借入金入力(wi01)
        
        GVar1 = wRs("実際年月日")
        If IsDate(GVar1) Then
            G借入金入力(wi01).借入返済年月日 = CDate(GVar1)
        End If
        GVar1 = wRs("利息計算年月日")
        If IsDate(GVar1) Then
            G借入金入力(wi01).利息計算年月日 = CDate(GVar1)
        End If
        
        G借入金入力(wi01).元金 = wRs("元金額")
        G借入金入力(wi01).利率 = wRs("利率")
        G借入金入力(wi01).利息額 = wRs("利息額")
        G借入金入力(wi01).仮計上利息額 = wRs("仮計上利息額")
        G借入金入力(wi01).返済金額 = wRs("返済金額")
        G借入金入力(wi01).融資残高 = wRs("融資残高")
        
        G借入金入力(wi01).日割日数 = wRs("日割日数")
        G借入金入力(wi01).利息対象期間日数 = wRs("利息対象期間日数")
        
        '利息額自動計算
        If p借入データ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
            G借入金入力(wi01).仮計上利息額 = MBD010_利息計算小数点5桁(G借入金入力(wi01).利率, _
                        G借入金入力(wi01).融資残高, G借入金入力(wi01).日割日数, p借入データ.金利計算年間日数)
        Else
            G借入金入力(wi01).仮計上利息額 = MBD010_利息計算小数点5桁(G借入金入力(wi01).利率, _
                        G借入金入力(wi01).元金 + G借入金入力(wi01).融資残高, G借入金入力(wi01).日割日数, p借入データ.金利計算年間日数)
        End If
        
        wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_借入金入力明細_利息額自動計算_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入金入力明細_利息額自動計算() でエラー" + vbCrLf + vbCrLf + _
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
' MBD010_借入利子補給金
'------------------------------------------------
Public Function MBD010_借入利子補給金(pNo As String) As Integer
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
'
    On Error GoTo MBD010_借入内入_ERR
'
    MBD010_借入利子補給金 = 0

    wstr = ""
    wstr = wstr & "Select K.借入番号 , S.利子補給金フラグ"
    wstr = wstr & " FROM DBDA010_借入金 As K"
    wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
    wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分"
    wstr = wstr & " Where K.借入番号='" & pNo & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.EOF Then
        MBD010_借入利子補給金 = P8.FCDbl(wRs("利子補給金フラグ"))
    End If
    
    wRs.Close
    Set wRs = Nothing
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_借入内入_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_借入利子補給金() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)
    
    End
'
End Function

