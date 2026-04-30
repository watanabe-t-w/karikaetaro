Attribute VB_Name = "MRB010_リース残高"
Option Explicit
'
Private Const pPROGRAM_ID As String = "MRB010_リース残高"
'
'------------------------------------------------
' MRB010_手入力リース残高表
'------------------------------------------------
Public Sub MRB010_手入力リース残高表(pTbl As String, pTbl2 As String)
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String, wWhere As String
    
    Dim pリース計画マスタ As MAA910_リース                                    '5/10/8 V129
    Dim wリースマスタ As MAA030_リース
    Dim wリース As MAA910_リース
    
    Dim j As Integer, k As Integer
    Dim w千円単位 As Integer
    Dim w年 As Integer, w月 As Integer, w日 As Integer                         '5/8/18 V129
    Dim ww月 As Integer                                                        '5/9/2 V129
    Dim w間隔 As Integer, w回数 As Integer, wcnt As Integer
    
    '5/9/12 V129 解約の時 1回前の年月との差＞＝２　1回前のリース総額残更新
    Dim w月差 As Integer
    Dim w処理 As Integer                                                       '5/9/9 V129
    
    Dim w残高算出 As Double                                                  '5/8/18 V129
    Dim wリース総額残高 As Double                                                  '5/9/9 V129
    Dim w前月残合計 As Double, w前月残高(12) As Double                          ' 07/02/09 V180
    Dim wリース総額合計 As Double, wリース総額(12) As Double
    Dim w消費税総額合計 As Double, w消費税総額(12) As Double
    Dim wリース料合計 As Double, wリース料(12) As Double
    Dim w消費税額合計 As Double, w消費税額(12) As Double
    Dim w支払合計額合計 As Double, w支払合計額(12) As Double
    Dim w解約合計 As Double, w解約(12) As Double
    Dim w支払残高合計 As Double, w支払残高(12) As Double
    Dim w保証合計 As Double, w保証(12) As Double
    
    
    Dim w基準年月 As Date                                   '2008/02/06 V182
    Dim w終了年月 As Variant                                '2008/02/06 V182
    Dim w開始回目 As Integer                                '2008/02/07 V182
    Dim w終了回目 As Integer                                '2008/02/07 V182
    Dim w回目 As Integer                                    '2008/02/06 V182
    Dim w解約締切年月日 As Variant                          '2008/02/06 V182
    
    Dim w利息対象期間日数 As Integer                        '2008/02/07 V182
    Dim w利息区分 As String                                 '2008/02/07 V182
    
        
    Dim wd01 As Date
    Dim w実際年月 As Date
    Dim w開始年月日 As Date
    Dim w年月(12) As Date
    Dim w実際年月日OLD As Date                                                 '5/8/18 V129
    Dim w対象年月 As Date, w対象年月調整 As Date                               'V120
    Dim w対象年月OLD As Date, w対象年月NEW As Date                             '5/8/18 V129
    
    Dim w解約実行日 As Variant                                                 '5/10/8 V129
    Dim w管理年月1 As Variant, w管理年月2 As Variant, w管理年月3 As Variant    '5/9/8 V129
    Dim w実績年月1 As Variant, w実績年月2 As Variant, w実績年月3 As Variant    '5/9/8 V129
    Dim w実績年月日1 As Variant, w実績年月日2 As Variant, w実績年月日3 As Variant '5/9/8 V129
    Dim w集計年月 As Variant                                                   '5/10/8 V129
    
    Dim w分母 As String
    Dim w開始年 As String
    Dim wリース番号 As String, wリース計画番号 As String
    
    Dim w解約判定 As String
    Dim ws基本 As String                                                       '5/8/30 V129
    Dim wベンチャ As String, wベンチャcode As String
    
    Dim w前月残 As Double                                                       ' 07/02/09 V180
    
    
    
    '5/10/9 V129 支店への貸付　リストラ番号
    Dim w期首リース年度 As String, w支店貸付 As String, w全社リース As String
'    Dim wsTbl As String
'
    On Error GoTo MRB010_手入力リース残高表_ERR
'
    ' -----------------------------------------
    '       リース金マスタより DCDA010_リース残高推移表結果　作成
    ' -----------------------------------------
    w開始年 = GRpt.テキスト_01
    wリース計画番号 = GRpt.リス
    'w金融リストラ = GRpt.金融
    w千円単位 = GRpt.チェック_01

    If w千円単位 = 1 Then
        w分母 = "1000"
    Else
        w分母 = "1"
    End If
    
    wベンチャ = Left$(wリース計画番号, 4)
    wベンチャcode = StrConv(Right$(wベンチャ, 1), 2)

    ws基本 = Left$(wリース計画番号, 2)            '5/8/30 V129
    If ws基本 = wリース計画番号 Then
        wベンチャcode = "a"                     '5/8/30 V129
    End If                                      '5/8/30 V129
'
    '**年月テーブル作成
    Select Case GRpt.推移
        Case "月次"
            w間隔 = 1: wcnt = 12
        Case "四半期"
            w間隔 = 3: wcnt = 12
        Case "半期"
            w間隔 = 6: wcnt = 10
        Case "年次"
            w間隔 = 12: wcnt = 10
    End Select
'
    'w開始年月日 = C年月日.年度開始年月日(w開始年, "平成")
    '2019/01/15 日付入力区分 仕様変更
    If G基本情報.日付入力区分 = "0" Then
    '和暦
        If Len(w開始年) <= 2 Then
            w開始年月日 = C年月日.年度開始年月日(w開始年, "平成")
        Else
            w開始年月日 = C年月日.年度開始年月日(w開始年, "西暦")
        End If
    Else
    '西暦
        w開始年月日 = C年月日.年度開始年月日(w開始年, "西暦")
    End If
    
    w開始年月日 = DateAdd("m", -1, w開始年月日)
    
    w年月(0) = w開始年月日                          '5/10/8 V129
'
    For j = 1 To 12                                 '5/10/8 V129
        w年月(j) = DateAdd("m", w間隔, w年月(j - 1))
    Next
'
    If wリース計画番号 = "" Then                     '5/10/17 V129
        w期首リース年度 = w開始年                    '5/10/17 V129
    Else                                           '5/10/17 V129
        w期首リース年度 = Left$(wリース計画番号, 2)    '5/10/9 V129
    End If                                         '5/10/17 V129
    
    w支店貸付 = "支店貸付"                         '5/10/9 V129
    w全社リース = "全社リース"                         '5/10/17 V129
'
    wstr2 = ""
    wstr2 = wstr2 + "Select * From DCDA010_リース残高推移表結果"
    Call AdoRecordsetOpen(GDb, wRs2, wstr2)
    
    wstr = ""
    wstr = wstr + "Select * From " & pTbl
    wstr = wstr + " Where 手入力区分 = 1"                   ' 07/02/09 V180
    
    '**ベンチャーの場合
    If (wベンチャcode < "a" Or wベンチャcode > "z") And wリース計画番号 <> "" Then
        wstr = wstr + " And ("                               ' 07/02/09 V180
        
        wstr = wstr + " (リース計画番号='" & wリース計画番号 & "' And リース計画番号 <> '" & "'"
        
        If G基本情報.企業区分 = "連結本部" Then         '06/09/22 V170
            wstr = wstr + " And Format(実行日,'yyyymmdd') > '" & Format(Gコントロール.最終実績年月, "yyyymmdd") & "'"
        End If                                    '17/8/17 V128
        
        wstr = wstr + " And 取消フラグ=0)"
    
    Else
    '**会社全体の場合
        wstr = wstr + " And ((sm区分=0 And 取消フラグ=0)"
        
        If G基本情報.企業区分 = "本部" Or G基本情報.企業区分 = "連結本部" Then      '06/09/29 V170                          '06/05/02 V150
'                wstr = wstr + " Or (リース計画番号 = '" & pリース計画番号 & "'"  '06/05/02 V150
'                wstr = wstr + " And 取消フラグ = 0)"                         '06/05/02 V150
        
            wstr = wstr + " Or (リース計画番号='' And sm区分 = 0 And 取消フラグ = 0)"
            wstr = wstr + " Or (リース計画番号 = '" & wリース計画番号 & "' And リース計画番号 <> ''"
            wstr = wstr + " And sm区分 = 0 And 取消フラグ = 0)"
        
        End If                                                          '06/05/02 V150
         
         wstr = wstr + " Or (リース計画番号='" & wリース計画番号 & "' And リース計画番号 <> '" & "'"
        
        If G基本情報.企業区分 = "連結本部" Then         '06/09/22 V170
            wstr = wstr + " And Format(実行日,'yyyymmdd') > '" & Format(Gコントロール.最終実績年月, "yyyymmdd") & "'"
        End If                                    '17/8/17 V128
        
        wstr = wstr + " And 取消フラグ=0)"
        
        If G基本情報.企業区分 = "連結本部" Or G基本情報.企業区分 = "連結子会社" Or G基本情報.企業区分 = "全社" Then
            wstr = wstr + " Or (リース計画番号='" & wリース計画番号 & "' And sm区分=1 And 取消フラグ=0)"
        End If
    
    End If
        wstr = wstr + ")"

    If pTbl2 <> "" Then
        wstr = wstr + " UNION Select * From " & pTbl2
        wstr = wstr + " Where 手入力区分 = 1"
        wstr = wstr + " And ((リース計画番号='' And sm区分 = 0 And 取消フラグ = 0)"
        wstr = wstr + " Or (リース計画番号 = '" & wリース計画番号 & "' And リース計画番号 <> ''"
        wstr = wstr + " And sm区分 = 0 And 取消フラグ = 0))"
    End If
        
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            pリース計画マスタ = MBA010_リースデータセット(wRs)      '5/10/8 V129
         
            '** リース金テーブル セット **
            'Call MBD010_リース金テーブル作成(w金融リストラ, pリース計画マスタ)
            
            wリース番号 = pリース計画マスタ.リース番号                '5/10/8 V129
            For j = 1 To 12
                w前月残高(j) = 0                                ' 07/02/09 V180
                wリース総額(j) = 0
                w消費税総額(j) = 0
                wリース料(j) = 0
                w消費税額(j) = 0
                w支払合計額(j) = 0
                w解約(j) = 0
                w支払残高(j) = 0
                w保証(j) = 0
                
                
            Next
            
            wリース総額合計 = 0
            w消費税総額合計 = 0
            wリース料合計 = 0
            w消費税額合計 = 0
            w支払合計額合計 = 0
            w解約合計 = 0
            w支払残高合計 = 0
            
            '
            
            ' =========================================
            '                 初期設定
            ' =========================================
            '期首前残設定
            
            '***
             Call MBA010_リース入力明細Read(pリース計画マスタ.リース番号)  ' 07/02/09 V180
            
            w対象年月 = DateAdd("m", 1, w年月(0))                   ' 07/02/09 V180
            w前月残 = MBA010_リース手入力残高(pリース計画マスタ, 0, w対象年月) ' 07/02/09 V180
            w前月残高(1) = w前月残                                  ' 07/02/09 V180
            w対象年月 = MBA010_対象年月(CDate(pリース計画マスタ.実行日))  ' 07/02/09 V180
            
            w基準年月 = MBA010_対象年月(CDate(pリース計画マスタ.実行日))      '2008/02/06 V182
            
            
              
             For k = 1 To wcnt                                                               '5/10/8 V129
                 If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                     And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then '5/10/8 V129
                     wリース総額(k) = wリース総額(k) + pリース計画マスタ.リース総額          '5/10/8 V129
                     w消費税総額(k) = w消費税総額(k) + pリース計画マスタ.消費税総額
                     Exit For                                                                '5/10/9 V129
                 End If                                                                      '5/10/8 V129
             Next                                                                            '5/10/8 V129
                 
                
             For j = 1 To UBound(Gリース入力)                       ' 07/02/09 V180
                w対象年月 = MBA010_対象年月(CDate(Gリース入力(j).リース支払年月日))
                For k = 1 To wcnt
                    If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                       And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then
                        wリース料(k) = wリース料(k) + Gリース入力(j).リース料
                        w消費税額(k) = w消費税額(k) + Gリース入力(j).消費税額
                        w支払合計額(k) = w支払合計額(k) + Gリース入力(j).支払合計額
                        Exit For
                    End If
                Next
             Next
             
             
             
             
             '残高算出
             For k = 1 To wcnt
                If k > 1 Then
                    w前月残高(k) = w支払残高(k - 1)
                End If
                
                w支払残高(k) = w前月残高(k) + wリース総額(k) + w消費税総額(k) - wリース料(k) - w消費税額(k)
             Next
             
          

               
             For k = 1 To wcnt
                wリース総額合計 = wリース総額合計 + wリース総額(k)
                w消費税総額合計 = w消費税総額合計 + w消費税総額(k)
                wリース料合計 = wリース料合計 + wリース料(k)
                w消費税額合計 = w消費税額合計 + w消費税額(k)
                w支払合計額合計 = w支払合計額合計 + w支払合計額(k)
                w解約合計 = w解約合計 + w解約(k)
                
             Next
             w支払残高合計 = w支払残高(wcnt)
             
                
             If wリース総額合計 = 0 And w消費税総額合計 = 0 And wリース料合計 = 0 And w消費税額合計 = 0 And _
                w支払合計額合計 = 0 And w解約合計 = 0 And _
                w支払残高合計 = 0 Then
             Else
                wRs2.AddNew
                    wRs2("リース番号") = wリース番号
                    wRs2("リース総額合計") = wリース総額合計
                    wRs2("消費税総額合計") = w消費税総額合計
                    wRs2("リース料合計") = wリース料合計
                    wRs2("消費税額合計") = w消費税額合計
                    wRs2("支払合計額合計") = w支払合計額合計
                    wRs2("解約合計") = w解約合計
                    wRs2("支払残高合計") = w支払残高合計
                    
                     
                     
                    For k = 1 To wcnt
                        wRs2("リース総額_" + CStr(Format(k, "00"))) = wリース総額(k)
                        wRs2("消費税総額_" + CStr(Format(k, "00"))) = w消費税総額(k)
                        wRs2("リース料_" + CStr(Format(k, "00"))) = wリース料(k)
                        wRs2("消費税額_" + CStr(Format(k, "00"))) = w消費税額(k)
                        wRs2("支払合計額_" + CStr(Format(k, "00"))) = w支払合計額(k)
                        wRs2("解約_" + CStr(Format(k, "00"))) = w解約(k)
                        wRs2("支払残高_" + CStr(Format(k, "00"))) = w支払残高(k)
                        
                        
                    Next
                        
                wRs2.Update
             End If
                
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    wRs2.Close
    Set wRs2 = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MRB010_手入力リース残高表_ERR:
    pERR_MES = pPROGRAM_ID + "/ MRB010_手入力リース残高表() でエラー" + vbCrLf + vbCrLf + _
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
' MRB010_標準入力リース残高表
'------------------------------------------------
Public Sub MRB010_標準入力リース残高表(pTbl As String, pTbl2 As String)
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String, wWhere As String
    
    Dim pリース計画マスタ As MAA910_リース                                       '5/10/8 V129
    Dim wリースマスタ As MAA030_リース
    Dim wリース As MAA910_リース
    
    Dim j As Integer, k As Integer
    Dim w千円単位 As Integer
    Dim w年 As Integer, w月 As Integer, w日 As Integer                         '5/8/18 V129
    Dim ww月 As Integer                                                        '5/9/2 V129
    Dim w間隔 As Integer, w回数 As Integer, wcnt As Integer
    
    '5/9/12 V129 解約の時 1回前の年月との差＞＝２　1回前のリース総額残更新
    Dim w月差 As Integer
    Dim w処理 As Integer                                                       '5/9/9 V129
    
    Dim w残高算出 As Double                                                  '5/8/18 V129
    Dim wリース総額残高 As Double                                                  '5/9/9 V129
    Dim w前月残合計 As Double, w前月残高(12) As Double                          ' 07/02/09 V180
    Dim wリース総額合計 As Double, wリース総額(12) As Double
    Dim w消費税総額合計 As Double, w消費税総額(12) As Double
    Dim wリース料合計 As Double, wリース料(12) As Double
    Dim w消費税額合計 As Double, w消費税額(12) As Double
    Dim w支払合計額合計 As Double, w支払合計額(12) As Double
    Dim w解約合計 As Double, w解約(12) As Double
    Dim w支払残高合計 As Double, w支払残高(12) As Double
    
    Dim w基準年月 As Date                                   '2008/02/06 V182
    Dim w終了年月 As Variant                                '2008/02/06 V182
    Dim w開始回目 As Integer                                '2008/02/07 V182
    Dim w終了回目 As Integer                                '2008/02/07 V182
    Dim w回目 As Integer                                    '2008/02/06 V182
    Dim w解約締切年月日 As Variant                          '2008/02/06 V182
    
    
        
    Dim wd01 As Date
    Dim w実際年月 As Date
    Dim w開始年月日 As Date
    Dim w年月(12) As Date
    Dim w実際年月日OLD As Date                                                 '5/8/18 V129
    Dim w対象年月 As Date, w対象年月調整 As Date                               'V120
    Dim w対象年月OLD As Date, w対象年月NEW As Date                             '5/8/18 V129
    
    Dim w解約実行日 As Variant                                                 '5/10/8 V129
    Dim w管理年月1 As Variant, w管理年月2 As Variant, w管理年月3 As Variant    '5/9/8 V129
    Dim w実績年月1 As Variant, w実績年月2 As Variant, w実績年月3 As Variant    '5/9/8 V129
    Dim w実績年月日1 As Variant, w実績年月日2 As Variant, w実績年月日3 As Variant '5/9/8 V129
    Dim w集計年月 As Variant                                                   '5/10/8 V129
    
    Dim w分母 As String
    Dim w開始年 As String
    Dim wリース番号 As String, wリース計画番号 As String
    
    Dim w解約判定 As String
    Dim ws基本 As String                                                       '5/8/30 V129
    Dim wベンチャ As String, wベンチャcode As String
    
    Dim w前月残 As Double                                                       ' 07/02/09 V180
    
    '5/10/9 V129 支店への貸付　リストラ番号
    Dim w期首リース年度 As String, w全社リース As String
'    Dim wsTbl As String
'
    On Error GoTo MRB010_標準入力リース残高表_ERR
'
    ' -----------------------------------------
    '       リースマスタより DCDA010_リース残高推移表結果　作成
    ' -----------------------------------------
    w開始年 = GRpt.テキスト_01
    wリース計画番号 = GRpt.リス
    w千円単位 = GRpt.チェック_01

    If w千円単位 = 1 Then
        w分母 = "1000"
    Else
        w分母 = "1"
    End If
    
    wベンチャ = Left$(wリース計画番号, 4)
    wベンチャcode = StrConv(Right$(wベンチャ, 1), 2)

    ws基本 = Left$(wリース計画番号, 2)            '5/8/30 V129
    If ws基本 = wリース計画番号 Then
        wベンチャcode = "a"                     '5/8/30 V129
    End If                                      '5/8/30 V129
'
    '**年月テーブル作成
    Select Case GRpt.推移
        Case "月次"
            w間隔 = 1: wcnt = 12
        Case "四半期"
            w間隔 = 3: wcnt = 12
        Case "半期"
            w間隔 = 6: wcnt = 10
        Case "年次"
            w間隔 = 12: wcnt = 10
    End Select
'
    'w開始年月日 = C年月日.年度開始年月日(w開始年, "平成")
    '2019/01/15 日付入力区分 仕様変更
    If G基本情報.日付入力区分 = "0" Then
    '和暦
        If Len(w開始年) <= 2 Then
            w開始年月日 = C年月日.年度開始年月日(w開始年, "平成")
        Else
            w開始年月日 = C年月日.年度開始年月日(w開始年, "西暦")
        End If
    Else
    '西暦
        w開始年月日 = C年月日.年度開始年月日(w開始年, "西暦")
    End If
    
    w開始年月日 = DateAdd("m", -1, w開始年月日)
    
    w年月(0) = w開始年月日                          '5/10/8 V129
'
    For j = 1 To 12                                 '5/10/8 V129
        w年月(j) = DateAdd("m", w間隔, w年月(j - 1))
    Next
'
    If wリース計画番号 = "" Then                     '5/10/17 V129
        w期首リース年度 = w開始年                    '5/10/17 V129
    Else                                           '5/10/17 V129
        w期首リース年度 = Left$(wリース計画番号, 2)    '5/10/9 V129
    End If                                         '5/10/17 V129
    
    w全社リース = "全社リース"                         '5/10/17 V129
    
    
    '** ワークファイル 削除 **
    wstr2 = ""
    wstr2 = wstr2 + "Delete * From DCDA010_リース残高推移表結果"
    GDb.Execute wstr2
'
    wstr2 = ""
    wstr2 = wstr2 + "Select * From DCDA010_リース残高推移表結果"
    Call AdoRecordsetOpen(GDb, wRs2, wstr2)
'
    wstr = ""
    wstr = wstr + "Select * From " & pTbl
    wstr = wstr + " Where 手入力区分 = 0"                   ' 07/02/09 V180
    
    '**ベンチャーの場合
    If (wベンチャcode < "a" Or wベンチャcode > "z") And wリース計画番号 <> "" Then
        wstr = wstr + " And ("                               ' 07/02/09 V180
        
         
        wstr = wstr + " (リース計画番号='" & wリース計画番号 & "' And リース計画番号 <> '" & "'"
        
        If G基本情報.企業区分 = "連結本部" Then         '06/09/22 V170
            wstr = wstr + " And Format(実行日,'yyyymmdd') > '" & Format(Gコントロール.最終実績年月, "yyyymmdd") & "'"
        End If                                    '17/8/17 V128
        
        wstr = wstr + " And 取消フラグ=0)"
    
    Else
    '**会社全体の場合
        wstr = wstr + " And ((sm区分=0 And 取消フラグ=0)"
        
        If G基本情報.企業区分 = "本部" Or G基本情報.企業区分 = "連結本部" Then      '06/09/29 V170                          '06/05/02 V150
        
            wstr = wstr + " Or (リース計画番号='' And sm区分 = 0 And 取消フラグ = 0)"
            wstr = wstr + " Or (リース計画番号 = '" & wリース計画番号 & "' And リース計画番号 <> ''"
            wstr = wstr + " And sm区分 = 0 And 取消フラグ = 0)"
        
        End If                                                          '06/05/02 V150
            
        
        wstr = wstr + " Or (リース計画番号='" & wリース計画番号 & "' And リース計画番号 <> '" & "'"
        
        If G基本情報.企業区分 = "連結本部" Then         '06/09/22 V170
            wstr = wstr + " And Format(実行日,'yyyymmdd') > '" & Format(Gコントロール.最終実績年月, "yyyymmdd") & "'"
        End If                                    '17/8/17 V128
        
        wstr = wstr + " And 取消フラグ=0)"
        
        If G基本情報.企業区分 = "連結本部" Or G基本情報.企業区分 = "連結子会社" Or G基本情報.企業区分 = "全社" Then
            wstr = wstr + " Or (リース計画番号='" & wリース計画番号 & "' And sm区分=1 And 取消フラグ=0)"
        End If
    
    End If
        wstr = wstr + ")"

    If pTbl2 <> "" Then
        wstr = wstr + " UNION Select * From " & pTbl2
        wstr = wstr + " Where 手入力区分 = 1"
        wstr = wstr + " And ((リース計画番号='' And sm区分 = 0 And 取消フラグ = 0)"
        wstr = wstr + " Or (リース計画番号 = '" & wリース計画番号 & "' And リース計画番号 <> ''"
        wstr = wstr + " And sm区分 = 0 And 取消フラグ = 0))"
    End If
        
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            pリース計画マスタ = MBA010_リースデータセット(wRs)      '5/10/8 V129
         
            '** リーステーブル セット **
            
            Call MBA010_リーステーブル作成(pリース計画マスタ)    ' 07/02/18 V180
            
            wリース番号 = pリース計画マスタ.リース番号                '5/10/8 V129
            For j = 1 To 12
                w前月残高(j) = 0                                ' 07/02/09 V180
                wリース総額(j) = 0
                w消費税総額(j) = 0
                wリース料(j) = 0
                w消費税額(j) = 0
                w支払合計額(j) = 0
                w解約(j) = 0
                w支払残高(j) = 0
                
            Next
            
            wリース総額合計 = 0
            w消費税総額合計 = 0
            wリース料合計 = 0
            w消費税額合計 = 0
            w支払合計額合計 = 0
            w解約合計 = 0
            w支払残高合計 = 0
             
            
            '
            
            ' =========================================
            '                 初期設定
            ' =========================================
            '期首前残設定
            
            '***
            
            
            w対象年月 = DateAdd("m", 1, w年月(0))                   ' 07/02/09 V180
            'w前月残 = MBD010_リース金手入力残高(pリース計画マスタ, 0, w対象年月) ' 07/02/09 V180
            w前月残 = MBA010_リース標準入力残高(pリース計画マスタ, 0, w対象年月, G基本情報.借入金管理区分) '07/02/26 V180
            w前月残高(1) = w前月残                                  ' 07/02/09 V180
            w対象年月 = MBA010_対象年月(CDate(pリース計画マスタ.実行日))  ' 07/02/09 V180
            
            w基準年月 = MBA010_対象年月(CDate(pリース計画マスタ.実行日))      '2008/02/06 V182
            
              
             For k = 1 To wcnt                                                               '5/10/8 V129
                 If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                     And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then '5/10/8 V129
                     wリース総額(k) = wリース総額(k) + pリース計画マスタ.リース総額         '5/10/8 V129
                     w消費税総額(k) = w消費税総額(k) + pリース計画マスタ.消費税総額
                     Exit For                                                                '5/10/9 V129
                 End If                                                                      '5/10/8 V129
             Next                                                                            '5/10/8 V129
                 
             For j = 1 To UBound(Gリーステーブル)                   ' 07/02/18 V180
             
                Call MBA010_借入金年月算出(Gリーステーブル(j).支払予定年月, _
                    Gリーステーブル(j).実際年月日, pリース計画マスタ.支払日)  ' 07/02/12 V180
                    
                If G基本情報.借入金管理区分 = XMXA020_区分("借入金管理区分", "管理用") Then '07/02/18 V180
                    w対象年月 = MBA010_対象年月((CDate(G管理年月)))             ' 07/02/18 V180
                Else                                                            ' 07/02/180V180
                    w対象年月 = MBA010_対象年月((CDate(G実績年月)))             ' 07/02/18 V180
                End If                                                          ' 07/02/180V180
                
                '解約算出
                w解約実行日 = pリース計画マスタ.解約実行日
               
                
                If Format(w解約実行日, "yyyymmdd") = _
                            Format(Gリーステーブル(j).実際年月日, "yyyymmdd") Then  ' 07/02/18 V180
                    w対象年月 = MBA010_対象年月((CDate(G実績年月)))             ' 07/02/18 V180
                End If                                                          ' 07/02/18 V180
                        
             
                For k = 1 To wcnt
                    If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                       And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then
                        wリース料(k) = wリース料(k) + Gリーステーブル(j).リース料     ' 07/02/18 V180
                        w消費税額(k) = w消費税額(k) + Gリーステーブル(j).消費税額     ' 07/02/18 V180
                        w支払合計額(k) = w支払合計額(k) + Gリーステーブル(j).支払合計額   ' 07/02/18 V180
                        'w保証(K) = w保証(K) + Gリーステーブル(J).保証料     ' 07/02/18 V180
                        
                        If Format(w解約実行日, "yyyymmdd") = _
                            Format(Gリーステーブル(j).実際年月日, "yyyymmdd") Then  ' 07/02/18 V180
                            w解約(k) = w解約(k) + Gリーステーブル(j).支払残高       ' 07/02/18 V180
                        End If                                                  ' 07/02/18 V180
                        
                        Exit For
                    End If
                Next
             Next
             
             
             
             For k = 1 To wcnt
                If k > 1 Then
                    w前月残高(k) = w支払残高(k - 1)
                End If
                
                w支払残高(k) = w前月残高(k) + wリース総額(k) + w消費税総額(k) - wリース料(k) - w消費税額(k) - w解約(k) ' 07/02/18 V180
                
             Next
             
          

               
             For k = 1 To wcnt
                wリース総額合計 = wリース総額合計 + wリース総額(k)
                w消費税総額合計 = w消費税総額合計 + w消費税総額(k)
                wリース料合計 = wリース料合計 + wリース料(k)
                w消費税額合計 = w消費税額合計 + w消費税額(k)
                w支払合計額合計 = w支払合計額合計 + w支払合計額(k)
                w解約合計 = w解約合計 + w解約(k)
                
             Next
             w支払残高合計 = w支払残高(wcnt)
                 
             If wリース総額合計 = 0 And w消費税総額合計 = 0 And wリース料合計 = 0 And w消費税額合計 = 0 And _
                w支払合計額合計 = 0 And w解約合計 = 0 And _
                w支払残高合計 = 0 Then
             Else
                wRs2.AddNew
                    wRs2("リース番号") = wリース番号
                    wRs2("リース総額合計") = wリース総額合計
                    wRs2("消費税総額合計") = w消費税総額合計
                    wRs2("リース料合計") = wリース料合計
                    wRs2("消費税額合計") = w消費税額合計
                    wRs2("支払合計額合計") = w支払合計額合計
                    wRs2("解約合計") = w解約合計
                    'wRs2("保証合計") = w保証合計
                    wRs2("支払残高合計") = w支払残高合計
                    
                      
                    For k = 1 To wcnt
                        wRs2("リース総額_" + CStr(Format(k, "00"))) = wリース総額(k)
                        wRs2("消費税総額_" + CStr(Format(k, "00"))) = w消費税総額(k)
                        wRs2("リース料_" + CStr(Format(k, "00"))) = wリース料(k)
                        wRs2("消費税額_" + CStr(Format(k, "00"))) = w消費税額(k)
                        wRs2("支払合計額_" + CStr(Format(k, "00"))) = w支払合計額(k)
                        wRs2("解約_" + CStr(Format(k, "00"))) = w解約(k)
                        wRs2("支払残高_" + CStr(Format(k, "00"))) = w支払残高(k)
                        
                    Next
                        
                wRs2.Update
             End If
                
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    wRs2.Close
    Set wRs2 = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MRB010_標準入力リース残高表_ERR:
    pERR_MES = pPROGRAM_ID + "/ MRB010_標準入力リース残高表() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

