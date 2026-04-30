Attribute VB_Name = "MDA020_社債他"
Option Explicit
'
Private Const pPROGRAM_ID As String = "MDA020_社債他"
'
'------------------------------------------------
' MDA020_借入金入力明細社債Read
'------------------------------------------------
Public Sub MDA020_借入金入力明細社債Read(p借入金 As MAA910_借入金)
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    Dim wi01 As Integer
'
    On Error GoTo MDA020_借入金入力明細社債Read_ERR
'
    ReDim G社債入力(0)
    wi01 = 0

    wstr = ""
    wstr = wstr & "Select"
    wstr = wstr & " 返済予定年月,実際年月日,初期手数料,元金手数料,利息手数料,保証料"
    wstr = wstr & " From DBDA010_借入金明細TR2"
    wstr = wstr & " Where 借入番号='" & p借入金.借入番号 & "'"
    wstr = wstr & " And 取消フラグ=0"
    wstr = wstr & " And 取消フラグ２=0"
    wstr = wstr & " Order by 実際年月日"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    Do Until wRs.eof
    
        wi01 = wi01 + 1
        ReDim Preserve G社債入力(wi01)
        
        GVar1 = wRs("実際年月日")
        If IsDate(GVar1) Then
            G社債入力(wi01).借入返済年月日 = CDate(GVar1)
        End If
        
        G社債入力(wi01).利息計算年月日 = G社債入力(wi01).借入返済年月日
        G社債入力(wi01).初期手数料 = wRs("初期手数料")
        G社債入力(wi01).元金手数料 = wRs("元金手数料")
        G社債入力(wi01).利息手数料 = wRs("利息手数料")
        G社債入力(wi01).保証料 = wRs("保証料")
        
        wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA020_借入金入力明細社債Read_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA020_借入金入力明細社債Read() でエラー" + vbCrLf + vbCrLf + _
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
' MDA020_借入金入力社債明細作成
'------------------------------------------------
Public Sub MDA020_借入金入力社債明細作成(p借入金 As MAA910_借入金)
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    
    Dim wi01 As Integer
    Dim j As Integer, k As Integer, l As Integer
    
    Dim w社債入力() As MAA910_借入金入力
    Dim w借入金入力() As MAA910_借入金入力
'
    On Error GoTo MDA020_借入金入力社債明細作成_ERR
'
    ReDim w社債入力(0)
    wi01 = 0

    wstr = ""
    wstr = wstr & "Select"
    wstr = wstr & " 返済予定年月,実際年月日,初期手数料,元金手数料,利息手数料,保証料"
    wstr = wstr & " From DBDA010_借入金明細TR2"
    wstr = wstr & " Where 借入番号='" & p借入金.借入番号 & "'"
    wstr = wstr & " And 取消フラグ=0"
    wstr = wstr & " And 取消フラグ２=0"
    wstr = wstr & " Order by 実際年月日"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    Do Until wRs.eof
    
        wi01 = wi01 + 1
        ReDim Preserve w社債入力(wi01)
        
        GVar1 = wRs("実際年月日")
        If IsDate(GVar1) Then
            w社債入力(wi01).借入返済年月日 = CDate(GVar1)
        End If
        
        w社債入力(wi01).利息計算年月日 = w社債入力(wi01).借入返済年月日
        w社債入力(wi01).初期手数料 = wRs("初期手数料")
        w社債入力(wi01).元金手数料 = wRs("元金手数料")
        w社債入力(wi01).利息手数料 = wRs("利息手数料")
        w社債入力(wi01).保証料 = wRs("保証料")
        
        wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
'
    wi01 = UBound(G借入金入力)
    ReDim w借入金入力(wi01)
    w借入金入力 = G借入金入力
'
    ReDim G借入金入力(0)
    wi01 = 0
    
    wstr = "SELECT 借入番号,実際年月日"
    wstr = wstr & " From DBDA010_借入金明細TR"
    wstr = wstr & " Where 借入番号 = '" & p借入金.借入番号 & "'"
    wstr = wstr & " ORDER BY 実際年月日"
    wstr = wstr & " union SELECT 借入番号,実際年月日"
    wstr = wstr & " From DBDA010_借入金明細TR2"
    wstr = wstr & " Where 借入番号 = '" & p借入金.借入番号 & "'"
    wstr = wstr & " ORDER BY 実際年月日"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    Do Until wRs.eof
    
        wi01 = wi01 + 1
        ReDim Preserve G借入金入力(wi01)
        
        GVar1 = wRs("実際年月日")
        If IsDate(GVar1) Then
            G借入金入力(wi01).借入返済年月日 = CDate(GVar1)
        End If
        
        wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
'
    For j = 1 To UBound(G借入金入力)
    
        For k = 1 To UBound(w社債入力)
            If G借入金入力(j).借入返済年月日 = w社債入力(k).借入返済年月日 Then
                G借入金入力(j).利息計算年月日 = w社債入力(k).利息計算年月日
                G借入金入力(j).初期手数料 = w社債入力(k).初期手数料
                G借入金入力(j).元金手数料 = w社債入力(k).元金手数料
                G借入金入力(j).利息手数料 = w社債入力(k).利息手数料
                G借入金入力(j).保証料 = w社債入力(k).保証料
                
                Exit For
            ElseIf G借入金入力(j).借入返済年月日 < w社債入力(k).借入返済年月日 Then
                Exit For
            End If
        Next k
        
        For l = 1 To UBound(w借入金入力)
            If G借入金入力(j).借入返済年月日 = w借入金入力(l).借入返済年月日 Then
                G借入金入力(j).利息計算年月日 = w借入金入力(l).利息計算年月日
                G借入金入力(j).元金 = w借入金入力(l).元金
                G借入金入力(j).利率 = w借入金入力(l).利率
                G借入金入力(j).利息額 = w借入金入力(l).利息額
                G借入金入力(j).仮計上利息額 = w借入金入力(l).仮計上利息額
                G借入金入力(j).返済金額 = w借入金入力(l).返済金額
                G借入金入力(j).融資残高 = w借入金入力(l).融資残高
                G借入金入力(j).日割日数 = w借入金入力(l).日割日数
                G借入金入力(j).利息対象期間日数 = w借入金入力(l).利息対象期間日数
                
                Exit For
            ElseIf G借入金入力(j).借入返済年月日 < w借入金入力(l).借入返済年月日 Then
                Exit For
            End If
        Next l
        
    Next j
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA020_借入金入力社債明細作成_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA020_借入金入力社債明細作成() でエラー" + vbCrLf + vbCrLf + _
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
' MDA020_社債借入残高
'------------------------------------------------
Public Sub MDA020_社債借入残高(pTbl As String)
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String, wWhere As String
    
    Dim p借入計画マスタ As MAA910_借入金                                       '5/10/8 V129
    Dim w借入金 As MAA910_借入金
    
    Dim j As Integer, k As Integer
    Dim w千円単位 As Integer
    Dim w間隔 As Integer, w回数 As Integer, wcnt As Integer
    
    Dim w保証合計 As Double, w保証(12) As Double
    Dim w初期手数料合計 As Double, w初期手数料(12) As Double                  '11/05/27 V190
    Dim w元金手数料合計 As Double, w元金手数料(12) As Double                  '11/05/27 V190
    Dim w利息手数料合計 As Double, w利息手数料(12) As Double                  '11/05/27 V190
    
    Dim w基準年月 As Date                                   '2008/02/06 V182
    Dim w開始年月日 As Date
    Dim w年月(12) As Date
    Dim w対象年月 As Date, w対象年月調整 As Date                               'V120
      
    Dim w分母 As String
    Dim w開始年 As String
    Dim w借入計画番号 As String, w金融リストラ As String
    Dim ws基本 As String                                                       '5/8/30 V129
    Dim wベンチャ As String, wベンチャcode As String
    '5/10/9 V129 支店への貸付　リストラ番号
    Dim w期首借入年度 As String, w支店貸付 As String, w全社借入 As String
'
    On Error GoTo MDA020_社債借入残高_ERR
'
    ' -----------------------------------------
    '       借入金マスタより DCDA010_借入残高推移表結果　作成
    ' -----------------------------------------
    w開始年 = GRpt.テキスト_01
    w借入計画番号 = GRpt.借入
    w金融リストラ = GRpt.金融
    w千円単位 = GRpt.千円単位

    If w千円単位 = 1 Then
        w分母 = "1000"
    Else
        w分母 = "1"
    End If
    
'    Select Case GRpt.帳票名
'    Case "借入残高推移表"
'        wsTbl = "DBDA010_借入金"
'    Case "貸付残高推移表"
'        wsTbl = "DBDA010_貸付金"
'    End Select
'
    wベンチャ = Left$(w借入計画番号, 4)
    wベンチャcode = StrConv(Right$(wベンチャ, 1), 2)

    ws基本 = Left$(w借入計画番号, 2)            '5/8/30 V129
    If ws基本 = w借入計画番号 Then
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
    If w借入計画番号 = "" Then                     '5/10/17 V129
        w期首借入年度 = w開始年                    '5/10/17 V129
    Else                                           '5/10/17 V129
        w期首借入年度 = Left$(w借入計画番号, 2)    '5/10/9 V129
    End If                                         '5/10/17 V129
    
    w支店貸付 = "支店貸付"                         '5/10/9 V129
    w全社借入 = "全社借入"                         '5/10/17 V129
'
    
    wstr = ""
    wstr = wstr + "Select * From " & pTbl
    wstr = wstr + " Where 手入力区分 = 1"                   ' 07/02/09 V180
    
    '**ベンチャーの場合
    If (wベンチャcode < "a" Or wベンチャcode > "z") And w借入計画番号 <> "" Then
        wstr = wstr + " And ("                               ' 07/02/09 V180
        
        If w金融リストラ <> "" Then
            wstr = wstr + " (金融リストラ番号='" & w金融リストラ & "' And 取消フラグ=0) Or  "
        End If
        
        wstr = wstr + " (借入計画番号='" & w借入計画番号 & "' And 借入計画番号 <> '" & "'"
        
        If G基本情報.企業区分 = "連結本部" Then         '06/09/22 V170
            wstr = wstr + " And Format(実行日,'yyyymmdd') > '" & Format(Gコントロール.最終実績年月, "yyyymmdd") & "'"
        End If                                    '17/8/17 V128
        
        wstr = wstr + " And 取消フラグ=0)"
    
    Else
    '**会社全体の場合
        wstr = wstr + " And ((sm区分=0 And 取消フラグ=0)"
        
        If G基本情報.企業区分 = "本部" Or G基本情報.企業区分 = "連結本部" Then      '06/09/29 V170                          '06/05/02 V150
'                wstr = wstr + " Or (借入計画番号 = '" & p借入計画番号 & "'"  '06/05/02 V150
'                wstr = wstr + " And 取消フラグ = 0)"                         '06/05/02 V150
        
            wstr = wstr + " Or (借入計画番号='' And sm区分 = 0 And 取消フラグ = 0)"
            wstr = wstr + " Or (借入計画番号 = '" & w借入計画番号 & "' And 借入計画番号 <> ''"
            wstr = wstr + " And sm区分 = 0 And 取消フラグ = 0)"
        
        End If                                                          '06/05/02 V150
            
        If w金融リストラ <> "" Then
            wstr = wstr + " Or (金融リストラ番号='" & w金融リストラ & "' And 取消フラグ=0)"
        End If
        
        wstr = wstr + " Or (金融リストラ番号='" & w期首借入年度 & "' And 取消フラグ=0)"
        wstr = wstr + " Or (金融リストラ番号='" & w支店貸付 & "' And 取消フラグ=0)"
        wstr = wstr + " Or (金融リストラ番号='" & w全社借入 & "' And 取消フラグ=0)"
        wstr = wstr + " Or (借入計画番号='" & w借入計画番号 & "' And 借入計画番号 <> '" & "'"
        
        If G基本情報.企業区分 = "連結本部" Then         '06/09/22 V170
            wstr = wstr + " And Format(実行日,'yyyymmdd') > '" & Format(Gコントロール.最終実績年月, "yyyymmdd") & "'"
        End If                                    '17/8/17 V128
        
        wstr = wstr + " And 取消フラグ=0)"
        
        If G基本情報.企業区分 = "連結本部" Or G基本情報.企業区分 = "連結子会社" Or G基本情報.企業区分 = "全社" Then
            wstr = wstr + " Or (借入計画番号='" & w借入計画番号 & "' And sm区分=1 And 取消フラグ=0)"
        End If
    
    End If
        wstr = wstr + ")"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            p借入計画マスタ = MBD010_借入データセット(wRs)      '5/10/8 V129
            
            For j = 1 To 12
                w保証(j) = 0
                 
                w初期手数料(j) = 0                  '11/05/27 V190
                w元金手数料(j) = 0                  '11/05/27 V190
                w利息手数料(j) = 0                  '11/05/27 V190
            Next
            
            w保証合計 = 0
            
            w初期手数料合計 = 0                     '11/05/27 V190
            w元金手数料合計 = 0                     '11/05/27 V190
            w利息手数料合計 = 0                     '11/05/27 V190
            
            '** 借入金テーブル セット **
             Call MDA020_借入金入力明細社債Read(p借入計画マスタ)  ' 07/02/09 V180
            
            w対象年月 = DateAdd("m", 1, w年月(0))                   ' 07/02/09 V180
            w対象年月 = MBA010_対象年月(CDate(p借入計画マスタ.実行日))  ' 07/02/09 V180
            w基準年月 = MBA010_対象年月(CDate(p借入計画マスタ.実行日))      '2008/02/06 V182
            
             For j = 1 To UBound(G社債入力)                       ' 07/02/09 V180
                w対象年月 = MBA010_対象年月(CDate(G社債入力(j).借入返済年月日))
                For k = 1 To wcnt
                    If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                       And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then
                        w保証(k) = w保証(k) + G社債入力(j).保証料
                        
                        w初期手数料(k) = w初期手数料(k) + G社債入力(j).初期手数料
                        w元金手数料(k) = w元金手数料(k) + G社債入力(j).元金手数料
                        w利息手数料(k) = w利息手数料(k) + G社債入力(j).利息手数料
                        Exit For
                    End If
                Next
             Next
             
             For k = 1 To wcnt
                w保証合計 = w保証合計 + w保証(k)
                
                w初期手数料合計 = w初期手数料合計 + w初期手数料(k)
                w元金手数料合計 = w元金手数料合計 + w元金手数料(k)
                w利息手数料合計 = w利息手数料合計 + w利息手数料(k)
             Next
              
             '***** 社債手数料による　DCDA010_借入残高推移表結果　の　更新処理
             If w保証合計 <> 0 Or _
                w初期手数料合計 <> 0 Or _
                w元金手数料合計 <> 0 Or _
                w利息手数料合計 Then
                
                wstr2 = ""
                wstr2 = wstr2 + "Select * From DCDA010_借入残高推移表結果"
                wstr2 = wstr2 + " Where 借入番号='" & p借入計画マスタ.借入番号 & "'"
                Call AdoRecordsetOpen(GDb, wRs2, wstr2)
                If wRs2.eof Then
                    wRs2.AddNew
                    wRs2("借入番号") = p借入計画マスタ.借入番号
                End If
                
                    wRs2("保証合計") = w保証合計
                    
                    wRs2("初期手数料合計") = w初期手数料合計    '11/05/27 V190
                    wRs2("元金手数料合計") = w元金手数料合計    '11/05/27 V190
                    wRs2("利息手数料合計") = w利息手数料合計    '11/05/27 V190
                    
                    For k = 1 To wcnt
                        wRs2("保証_" + CStr(Format(k, "00"))) = w保証(k)
                        
                        wRs2("初期手数料_" + CStr(Format(k, "00"))) = w初期手数料(k)
                        wRs2("元金手数料_" + CStr(Format(k, "00"))) = w元金手数料(k)
                        wRs2("利息手数料_" + CStr(Format(k, "00"))) = w利息手数料(k)
                    Next
                    
                    wRs2.Update
                                
                wRs2.Close
                Set wRs2 = Nothing
                
             End If
             
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA020_社債借入残高_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA020_社債借入残高() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub
