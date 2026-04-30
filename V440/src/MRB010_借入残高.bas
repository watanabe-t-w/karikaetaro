Attribute VB_Name = "MRB010_借入残高"
Option Explicit
'
Private Const pPROGRAM_ID As String = "MRB010_借入残高"

Dim p前払利息増(600) As Double         '2008/02/06 V182
Dim p前払利息減(600) As Double         '2008/02/06 V182
Dim p前払利息残(600) As Double         '2008/02/06 V182
Dim p未払利息増(600) As Double         '2008/02/06 V182
Dim p未払利息減(600) As Double         '2008/02/06 V182
Dim p未払利息残(600) As Double         '2008/02/06 V182

Dim w利息未払前払 As MAA910_利息未払前払テーブル    '11/02/11
Dim w前払利息残 As Double                           '11/02/11
Dim w未払利息残 As Double                           '11/02/11

Dim w手入力区分 As Integer                          '11/02/16

Dim OLD開始日 As Date                   '2016/09/26
Dim NEW開始日 As Date                   '2016/09/26
Dim OLD終了日 As Date                   '2016/09/26
Dim NEW終了日 As Date                   '2016/09/26
Dim WT合計利息額 As Double              '2016/09/26
Dim WT合計日数 As Integer               '2016/09/26
Dim WT集計利息額 As Double              '2016/09/26
 




Type MRB010_借入金推移表
    借入番号 As String
    
    融資合計 As Double
    融資(12) As Double
    元金合計 As Double
    元金(12) As Double
    利息合計 As Double
    利息(12) As Double
    返済合計 As Double
    返済(12) As Double
    解約合計 As Double
    解約(12) As Double
    残高合計 As Double
    残高(12) As Double
    保証合計 As Double
    保証(12) As Double
    手数料合計 As Double
    手数料(12) As Double
    
    初期手数料合計 As Double        '11/05/27 V190
    初期手数料(12) As Double        '11/05/27 V190
    元金手数料合計 As Double        '11/05/27 V190
    元金手数料(12) As Double        '11/05/27 V190
    利息手数料合計 As Double        '11/05/27 V190
    利息手数料(12) As Double        '11/05/27 V190
    
    
    前払利息増合計 As Double
    前払利息増(12) As Double
    前払利息減合計 As Double
    前払利息減(12) As Double
    前払利息合計 As Double
    前払利息(12) As Double
    未払利息増合計 As Double
    未払利息増(12) As Double
    未払利息減合計 As Double
    未払利息減(12) As Double
    未払利息合計 As Double
    未払利息(12) As Double
    
    長短振替額合計 As Double    '16/01/25
    長短振替額(12) As Double    '16/01/25
    
    損益利息額合計 As Double    '16/03/24
    損益利息額(12) As Double    '16/03/24
    
    '借入金種別区分 As String    '16/03/24
    利子補給金フラグ As Integer  '16/03/29
    
    利率(12) As Double
    
End Type

Dim ssw As Integer
Dim w直前利率 As Double
Dim w現在利率 As Double


'------------------------------------------------
' MRB010_手入力借入残高表
'------------------------------------------------
Public Sub MRB010_手入力借入残高表(pTbl As String)
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset, wRs3 As ADODB.Recordset    '16/01/25
        
    Dim wstr As String, wstr2 As String, wStr3 As String, wWhere As String          '16/01/25
    
    Dim wiCnt As Integer
    Dim p推移(9999) As MRB010_借入金推移表
    Dim FLG_Mdata As Boolean
    
    Dim p借入計画マスタ As MAA910_借入金                                       '5/10/8 V129
    Dim w銀行マスタ As MAA030_銀行
    Dim w借入金 As MAA910_借入金
    
    Dim j As Integer, k As Integer
    Dim w千円単位 As Integer
    Dim w年 As Integer, w月 As Integer, w日 As Integer                         '5/8/18 V129
    Dim ww月 As Integer                                                        '5/9/2 V129
    Dim w間隔 As Integer, w回数 As Integer, wcnt As Integer
    
    '5/9/12 V129 解約の時 1回前の年月との差＞＝２　1回前の融資残更新
    Dim w月差 As Integer
    Dim w処理 As Integer                                                       '5/9/9 V129
    
    Dim w残高年月 As Date                                                       '16/01/25
    
    
    
    Dim w残高算出 As Double                                                  '5/8/18 V129
    Dim w融資残高 As Double                                                  '5/9/9 V129
    Dim w前月残合計 As Double, w前月残高(12) As Double                          ' 07/02/09 V180
    Dim w融資合計 As Double, w融資(12) As Double
    Dim w元金合計 As Double, w元金(12) As Double
    Dim w利息合計 As Double, w利息(12) As Double
    Dim w返済合計 As Double, w返済(12) As Double
    Dim w解約合計 As Double, w解約(12) As Double
    Dim w残高合計 As Double, w残高(12) As Double
    Dim w保証合計 As Double, w保証(12) As Double
    Dim w手数料合計 As Double, w手数料(12) As Double           ' 08/12/08 V189
    
    Dim w初期手数料合計 As Double, w初期手数料(12) As Double                  '11/05/27 V190
    Dim w元金手数料合計 As Double, w元金手数料(12) As Double                  '11/05/27 V190
    Dim w利息手数料合計 As Double, w利息手数料(12) As Double                  '11/05/27 V190
    
    Dim w前払利息増合計 As Double, w前払利息増(12)          '2008/02/06 V182
    Dim w前払利息減合計 As Double, w前払利息減(12)          '2008/02/06 V182
    Dim w前払利息合計 As Double, w前払利息(12)              '2008/02/06 V182
    Dim w未払利息増合計 As Double, w未払利息増(12)          '2008/02/06 V182
    Dim w未払利息減合計 As Double, w未払利息減(12)          '2008/02/06 V182
    Dim w未払利息合計 As Double, w未払利息(12)              '2008/02/06 V182
    
    Dim w長短振替額合計 As Double, w長短振替額(12)          '16/01/25
    
    
    Dim w損益利息額合計 As Double, w損益利息額(12)          '16/03/24
    Dim w利子補給金フラグ As Integer                        '16/03/24
    
    
    Dim w利率(12) As Double                                 '11/02/17
    Dim w基準年月 As Date                                   '2008/02/06 V182
    Dim w終了年月 As Variant                                '2008/02/06 V182
    Dim w開始回目 As Integer                                '2008/02/07 V182
    Dim w終了回目 As Integer                                '2008/02/07 V182
    Dim w回目 As Integer                                    '2008/02/06 V182
    Dim w解約締切年月日 As Variant                          '2008/02/06 V182
    Dim w利息計算年月日 As Variant                          '10/02/04
    
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
    Dim w借入番号 As String, w借入計画番号 As String, w金融リストラ As String
    
    Dim w解約判定 As String
    Dim ws基本 As String                                                       '5/8/30 V129
    Dim wベンチャ As String, wベンチャcode As String
    
    Dim w借入貸付 As String                                                     ' 07/02/09 V180
    Dim w前月残 As Double                                                       ' 07/02/09 V180
    
    '5/10/9 V129 支店への貸付　リストラ番号
    Dim w期首借入年度 As String, w支店貸付 As String, w全社借入 As String
'    Dim wsTbl As String
'
    On Error GoTo MRB010_手入力借入残高表_ERR
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
    FLG_Mdata = False '通常はデータ一括書込
    wiCnt = 0
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

    'If pTbl2 <> "" Then
    '    wstr = wstr + " UNION Select * From " & pTbl2
    '    wstr = wstr + " Where 手入力区分 = 1"
    '    wstr = wstr + " And ((借入計画番号='' And sm区分 = 0 And 取消フラグ = 0)"
    '    wstr = wstr + " Or (借入計画番号 = '" & w借入計画番号 & "' And 借入計画番号 <> ''"
    '    wstr = wstr + " And sm区分 = 0 And 取消フラグ = 0))"
    'End If
        
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.RecordCount >= 10000 Then
        FLG_Mdata = True
        
        wstr2 = ""
        wstr2 = wstr2 + "Select * From DCDA010_借入残高推移表結果"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
    End If
        
        Do Until wRs.eof
            p借入計画マスタ = MBD010_借入データセット(wRs)      '5/10/8 V129
         
            '** 借入金テーブル セット **
            'Call MBD010_借入金テーブル作成(w金融リストラ, p借入計画マスタ)
            
            w借入番号 = p借入計画マスタ.借入番号                '5/10/8 V129
            
            w利子補給金フラグ = p借入計画マスタ.利子補給金フラグ    '16/03/24 利子補給に伴う変更
            
            For j = 1 To 12
                w前月残高(j) = 0                                ' 07/02/09 V180
                w融資(j) = 0
                w元金(j) = 0
                w利息(j) = 0
                w返済(j) = 0
                w解約(j) = 0
                w残高(j) = 0
                w保証(j) = 0
                w手数料(j) = 0                      ' 08/12/08 V189
                
                w長短振替額(j) = 0                  '16/01/25
                
                w損益利息額(j) = 0                  '16/03/24
                
                w初期手数料(j) = 0                  '11/05/27 V190
                w元金手数料(j) = 0                  '11/05/27 V190
                w利息手数料(j) = 0                  '11/05/27 V190
                
                w前払利息増(j) = 0                  '2008/02/06 V182
                w前払利息減(j) = 0                  '2008/02/06 V182
                w前払利息(j) = 0                    '2008/02/06 V182
                w未払利息増(j) = 0                  '2008/02/06 V182
                w未払利息減(j) = 0                  '2008/02/06 V182
                w未払利息(j) = 0                    '2008/02/06 V182
                
                w利率(j) = 0                        '11/02/17
                
            Next
            
            w融資合計 = 0
            w元金合計 = 0
            w利息合計 = 0
            w返済合計 = 0
            w解約合計 = 0
            w残高合計 = 0
            w保証合計 = 0
            w手数料合計 = 0                         ' 08/12/08 V189
            
            w長短振替額合計 = 0                     '16/01/25
            
            w損益利息額合計 = 0                     '16/3/24
            
            w初期手数料合計 = 0                     '11/05/27 V190
            w元金手数料合計 = 0                     '11/05/27 V190
            w利息手数料合計 = 0                     '11/05/27 V190
            
            w前払利息増合計 = 0                     '2008/02/06 V182
            w前払利息減合計 = 0                     '2008/02/06 V182
            w前払利息合計 = 0                       '2008/02/06 V182
            w未払利息増合計 = 0                     '2008/02/06 V182
            w未払利息減合計 = 0                     '2008/02/06 V182
            w未払利息合計 = 0                       '2008/02/06 V182
            
            For w回目 = 1 To 600                    '2008/02/06 V182
                p前払利息増(w回目) = 0              '2008/02/06 V182
                p前払利息減(w回目) = 0              '2008/02/06 V182
                p前払利息残(w回目) = 0              '2008/02/06 V182
                p未払利息増(w回目) = 0              '2008/02/06 V182
                p未払利息減(w回目) = 0              '2008/02/06 V182
                p未払利息残(w回目) = 0              '2008/02/06 V182
            Next                                    '2008/02/06 V182
            
            '
            
            ' =========================================
            '                 初期設定
            ' =========================================
            '期首前残設定
            
            w手入力区分 = p借入計画マスタ.手入力区分        '11/02/16
            
            '***
             Call MBD010_借入金入力明細Read(p借入計画マスタ)  ' 07/02/09 V180
            
            w対象年月 = DateAdd("m", 1, w年月(0))                   ' 07/02/09 V180
            w前月残 = MBD010_借入金手入力残高(p借入計画マスタ, 0, w対象年月) ' 07/02/09 V180
            w前月残高(1) = w前月残                                  ' 07/02/09 V180
            w対象年月 = MBA010_対象年月(CDate(p借入計画マスタ.実行日))  ' 07/02/09 V180
            
            w基準年月 = MBA010_対象年月(CDate(p借入計画マスタ.実行日))      '2008/02/06 V182
            
            
              
             For k = 1 To wcnt                                                               '5/10/8 V129
                 If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                     And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then '5/10/8 V129
                     w融資(k) = w融資(k) + p借入計画マスタ.融資金額                             '5/10/8 V129
                     Exit For                                                                '5/10/9 V129
                 End If                                                                      '5/10/8 V129
             Next                                                                            '5/10/8 V129
                
             '***手打ち入力　解約年月日SET　利息前払            ’10/01/13
             '   w解約締切年月日 = MBA010_締日年月日(CDate(w解約実行日))     '2008/02/06 V182
               
             w解約実行日 = Null                                 '10/01/13
             w解約締切年月日 = Null                             '10/01/13
             
             For j = 1 To UBound(G借入金入力)                     '10/01/13
                If p借入計画マスタ.利息区分 = "1" _
                        And Format(G借入金入力(j).借入返済年月日, "yyyy/mm/dd") _
                            = Format(p借入計画マスタ.最終返済実行日, "yyyy/mm/dd") _
                        And G借入金入力(j).利息額 < 0 Then              '10/01/13
                    w解約実行日 = p借入計画マスタ.最終返済実行日        '10/01/13
                    w解約締切年月日 = MBA010_締日年月日(CDate(w解約実行日))     '10/01/13
                    Exit For                                            '10/01/13
                End If                                                  '10/01/13
             Next                                                       '10/01/13
                    
                
             
                
             ssw = 0
             
             For j = 1 To UBound(G借入金入力)                        ' 07/02/09 V180
                w対象年月 = MBA010_対象年月(CDate(G借入金入力(j).借入返済年月日))
                For k = 1 To wcnt
                    If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                       And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then
                        w元金(k) = w元金(k) + G借入金入力(j).元金
                        w利息(k) = w利息(k) + G借入金入力(j).利息額
                        w返済(k) = w返済(k) + G借入金入力(j).返済金額
                        w利率(k) = G借入金入力(j).利率               '11/02/17
                        
                        
                        '*直前利率設定
                        If ssw = 0 Then
                            If j = 1 And G借入金入力(j).利率 = 0 Then
                                w直前利率 = p借入計画マスタ.利率
                            End If
                            
                            If j >= 2 Then
                                If G借入金入力(j - 1).利率 = 0 Then
                                    w直前利率 = p借入計画マスタ.利率
                                Else
                                    w直前利率 = G借入金入力(j - 1).利率
                                End If
                            End If
                            
                            ssw = 1
                        End If
                            
                        Exit For
                    End If
                Next
             
             '*****************************************************************
             '    利息前払　利息未払　メインルーチン  2008/02/06 V182
             '*****************************************************************
               
               'If G借入金入力 (j).利息額 <> 0 And _          2009/12/25
               '   ((Format(G借入金入力 (j).利息計算年月日, "yyyymmdd") _
               '    <> Format(p借入計画マスタ.最終返済実行日, "yyyymmdd")) Or _
               '    (Format(G借入金入力 (j).利息計算年月日, "yyyymmdd") _
               '    = Format(p借入計画マスタ.最終返済実行日, "yyyymmdd") And _
               '    Format(G借入金入力 (1).利息計算年月日, "yyyymmdd") _
               '    <> Format(p借入計画マスタ.実行日, "yyyymmdd"))) Then
                   
                
                'If Format(G借入金入力 (1).利息計算年月日, "yyyymmdd") _
                '     = Format(p借入計画マスタ.実行日, "yyyymmdd") And _
                '     G借入金入力 (1).利息額 <> 0 Then    '2009/12/15 V182
               If G借入金入力(j).利息額 <> 0 Or G借入金入力(j).仮計上利息額 Then                     ' 09/12/25
                'If p借入計画マスタ.利息区分 = "1" Then              ' 09/12/25
                '    w利息区分 = "1"                 '2008/02/07 V182
                '    If j = 1 Then                   '2008/02/07 V182
                '        w利息対象期間日数 = DateDiff("D", G借入金入力 (j).利息計算年月日, _
                '                                          G借入金入力 (j + 1).利息計算年月日) + 1  '2008/02/07 V182
                '    Else                            '2008/02/07 V182
                '        w利息対象期間日数 = DateDiff("D", G借入金入力 (j).利息計算年月日, _
                '                                          G借入金入力 (j + 1).利息計算年月日) '2008/02/07 V182
                '    End If                          '2008/02/07 V182
                    
                'Else                                '2008/02/07 V182
                '    w利息区分 = "2"                 '2008/02/07 V182
                '    If j = 1 Then                   '2008/07/07 V182
                '        w利息対象期間日数 = DateDiff("D", p借入計画マスタ.実行日, _
                '                                          G借入金入力 (j).利息計算年月日) + 1 '2008/02/07 V182
                '    Else                            '2008/02/07 V182
                '        w利息対象期間日数 = DateDiff("D", G借入金入力 (j - 1).利息計算年月日, _
                '                                          G借入金入力 (j).利息計算年月日) '2008/02/07 V182
                '    End If                          '2008/02/07 V182
                'End If                              '2008/02/07 V182
                
                    
                w対象年月 = MBA010_対象年月(CDate(G借入金入力(j).借入返済年月日))    '2008/02/07 V182
                w回目 = DateDiff("M", w基準年月, w対象年月) + 1     '2008/02/06 V182
                
                'w解約実行日 = Null                  '2008/02/07 V182
                'w解約締切年月日 = Null              '2008/02/07 V182
                 
                '日数調整、利息調整
                If G借入金入力(j).利息対象期間日数 = 0 Then          '10/01/08
                    w利息対象期間日数 = G借入金入力(j).日割日数      '10/01/08
                Else                                                '10/01/08
                    w利息対象期間日数 = G借入金入力(j).利息対象期間日数  '10/01/08
                End If                                              '10/01/08
                
                
                
                If p借入計画マスタ.利息区分 = "1" Then              '10/01/09 V182
                    If w回目 > 0 And w回目 <= 600 Then
                        p前払利息増(w回目) = p前払利息増(w回目) + G借入金入力(j).利息額  '2008/02/06 V182
                    End If
                    
                    If Format(G借入金入力(j).借入返済年月日, "yyyy/mm/dd") _
                        = Format(p借入計画マスタ.実行日, "yyyy/mm/dd") Then
                        If p借入計画マスタ.利息控除区分 = 1 Or p借入計画マスタ.利息控除区分 = 3 Then '10/02/04
                            w利息計算年月日 = DateAdd("d", 1, G借入金入力(j).利息計算年月日)                 '10/02/04
                        Else                                                                                '10/02/04
                            w利息計算年月日 = G借入金入力(j).利息計算年月日
                        End If                      '10/02/04
                        
                    Else                            '10/02/04
                        w利息計算年月日 = DateAdd("d", 1, G借入金入力(j).利息計算年月日)  '10/02/04
                    End If
                    
                    If Format(G借入金入力(j).借入返済年月日, "yyyymmdd") _
                       = Format(w解約実行日, "yyyymmdd") Then               '10/02/04
                        w利息計算年月日 = G借入金入力(j).利息計算年月日      '10/02/04
                    End If                                                  '10/02/04
                       
                    
                    Call MRB010_前払利息減計算(p借入計画マスタ, _
                                               p借入計画マスタ.実行日, _
                                               w利息計算年月日, _
                                               w利息対象期間日数, _
                                               G借入金入力(j).仮計上利息額 + G借入金入力(j).利息額, _
                                               w解約実行日, _
                                               w解約締切年月日, _
                                               w基準年月)     '2010/01/08 V182
                    
                  
                Else                                                '2008/02/06 V182
                    If w回目 > 0 And w回目 <= 600 Then
                        p未払利息減(w回目) = p未払利息減(w回目) + G借入金入力(j).利息額  '2008/02/06 V182
                    End If
                    
                    If Format(G借入金入力(j).借入返済年月日, "yyyy/mm/dd") _
                        = Format(p借入計画マスタ.最終返済実行日, "yyyy/mm/dd") _
                        And (p借入計画マスタ.利息控除区分 = 2 Or p借入計画マスタ.利息控除区分 = 3) Then '10/02/04
                        w利息計算年月日 = DateAdd("D", -1, G借入金入力(j).利息計算年月日)                '10/02/04
                    Else                                                                                '10/02/04
                        w利息計算年月日 = G借入金入力(j).利息計算年月日                                  '10/02/04
                    End If                                                                              '10/02/04
                    
                    Call MRB010_未払利息増計算(p借入計画マスタ, _
                                               p借入計画マスタ.実行日, _
                                               w利息計算年月日, _
                                               w利息対象期間日数, _
                                               G借入金入力(j).仮計上利息額 + G借入金入力(j).利息額, _
                                               w解約実行日, _
                                               w解約締切年月日, _
                                               w基準年月)     '2010/01/08 V182
                    
                End If                                              '2008/02/06 V182
                
               End If                                           '09/12/26
               
               
             Next
             
             '***終了年月算出
             If Not IsNull(w解約実行日) Then                    '2008/02/07 V182
                w終了年月 = w解約実行日                         '2008/02/07 V182
             Else                                               '2008/02/07 V182
                w終了年月 = p借入計画マスタ.最終返済実行日      '2008/02/07 V182
             End If                                             '2008/02/07 V182
             
             w終了年月 = MBA010_対象年月(CDate(w終了年月))      '2008/02/07 V182
             
             w開始回目 = DateDiff("M", w基準年月, w基準年月) + 1 '2008/02/07 V182
             w終了回目 = DateDiff("M", w基準年月, w終了年月) + 1 '2008/02/07 V182
             
             '**************************************************************
             '     前払利息　未払利息　集計セット
             '**************************************************************
             For j = w開始回目 To w終了回目                     '2008/02/07 V182
                w対象年月 = DateAdd("M", j - 1, w基準年月)      '2008/02/07 V182
                If j = 1 Then                                   '2008/02/07 V182
                    p前払利息残(j) = p前払利息増(j) - p前払利息減(j) '2008/02/07V182
                    p未払利息残(j) = p未払利息増(j) - p未払利息減(j) '2008/02/07V182
                Else                                            '2008/02/07 V182
                    p前払利息残(j) = p前払利息残(j - 1) + p前払利息増(j) - p前払利息減(j) '2008/02/07 V182
                    p未払利息残(j) = p未払利息残(j - 1) + p未払利息増(j) - p未払利息減(j) '2008/02/07V182
                End If                                          '2008/02/07 V182
                
                For k = 1 To wcnt                               '2008/02/07 V182
                    If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                        And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then '2008/02/07 V182
                            
                        w前払利息増(k) = w前払利息増(k) + p前払利息増(j) '2008/02/07 V182
                        w前払利息減(k) = w前払利息減(k) + p前払利息減(j) '2008/02/07 V182
                        w未払利息増(k) = w未払利息増(k) + p未払利息増(j) '2008/02/07 V182
                        w未払利息減(k) = w未払利息減(k) + p未払利息減(j) '2008/02/07 V182
                        
                        If Format(w対象年月, "yyyymmdd") = Format(w年月(k), "yyyymmdd") Then '2008/02/07 V182
                            w前払利息(k) = p前払利息残(j)       '2008/02/07 V182
                            w未払利息(k) = p未払利息残(j)       '2008/02/07 V182
                        End If                                  '2008/02/07 V182
                    End If                                      '2008/02/07 V182
                Next                                            '2008/02/07 V182
             Next                                               '2008/02/07 V182
             
             
             
             '残高算出
             For k = 1 To wcnt
                If k > 1 Then
                    w前月残高(k) = w残高(k - 1)
                End If
                
                w残高(k) = w前月残高(k) + w融資(k) - w元金(k)
             Next
             
          

               
             For k = 1 To wcnt
                If w解約(k) <> 0 Then               '11/06/17 V200
                    w元金(k) = w元金(k) + w解約(k)  '11/06/17 V200
                    w返済(k) = w返済(k) + w解約(k)  '11/06/17 V200
                End If                              '11/06/17 V200
                
                w融資合計 = w融資合計 + w融資(k)
                w元金合計 = w元金合計 + w元金(k)
                w利息合計 = w利息合計 + w利息(k)
                w返済合計 = w返済合計 + w返済(k)
                w解約合計 = w解約合計 + w解約(k)
                w保証合計 = w保証合計 + w保証(k)
                w手数料合計 = w手数料合計 + w手数料(k)                  ' 08/12/08 V189
                
                w前払利息増合計 = w前払利息増合計 + w前払利息増(k)      '2008/02/07 V182
                w前払利息減合計 = w前払利息減合計 + w前払利息減(k)      '2008/02/07 V182
                w未払利息増合計 = w未払利息増合計 + w未払利息増(k)      '2008/02/07 V182
                w未払利息減合計 = w未払利息減合計 + w未払利息減(k)      '2008/02/07 V182
                
                If p借入計画マスタ.利息区分 = "1" Then                  '16/03/24
                    w損益利息額(k) = w前払利息減(k)                     '16/03/24
                Else                                                    '16/03/24
                    w損益利息額(k) = w未払利息増(k)                     '16/03/24
                End If                                                  '16/03/24
                
                w損益利息額合計 = w損益利息額合計 + w損益利息額(k)      '16/03/24
                
                
             Next
             w残高合計 = w残高(wcnt)
             
             w前払利息合計 = w前払利息(wcnt)                            '2008/02/07 V182
             w未払利息合計 = w未払利息(wcnt)                            '2008/02/07 V182
             
             
             '***利率= 0 の時　調整処理
             w現在利率 = w直前利率
             For k = 1 To wcnt
                If w融資(k) <> 0 Or w元金(k) <> 0 Or w利息(k) <> 0 _
                                 Or w解約(k) <> 0 Or w残高(k) <> 0 Then
                    If w利率(k) = 0 Then
                        w利率(k) = w現在利率
                    End If
                    
                    w現在利率 = w利率(k)
                End If
             Next
             
             '***決算用の利率調整 2012/02/24
             For k = 1 To wcnt
                If (w残高(k) <> 0 Or w元金(k) <> 0 Or w利息(k)) And w利率(k) = 0 Then
                    If k = 1 Then
                        w利率(k) = w利率(k + 1)
                    Else
                        w利率(k) = w利率(k - 1)
                    End If
                End If
             Next
             
             
                
             If w融資合計 = 0 And w元金合計 = 0 And w利息合計 = 0 And _
                w返済合計 = 0 And w解約合計 = 0 And w保証合計 = 0 And _
                w前払利息増合計 = 0 And w前払利息減合計 = 0 And _
                w未払利息増合計 = 0 And w未払利息減合計 = 0 And _
                w残高合計 = 0 Then
             Else
                 If FLG_Mdata = True Then
                    wRs2.AddNew
                        wRs2("借入番号") = w借入番号
                        wRs2("融資合計") = w融資合計
                        wRs2("元金合計") = w元金合計
                        wRs2("利息合計") = w利息合計
                        wRs2("返済合計") = w返済合計
                        wRs2("解約合計") = w解約合計
                        wRs2("保証合計") = w保証合計
                        wRs2("手数料合計") = w手数料合計            ' 08/12/09 V189
                        
                        wRs2("初期手数料合計") = w初期手数料合計    '11/05/27 V190
                        wRs2("元金手数料合計") = w元金手数料合計    '11/05/27 V190
                        wRs2("利息手数料合計") = w利息手数料合計    '11/05/27 V190
                        
                        wRs2("残高合計") = w残高合計
    
                        wRs2("前払利息増合計") = w前払利息増合計    '2008/02/07 V182
                        wRs2("前払利息減合計") = w前払利息減合計    '2008/02/07 V182
                        wRs2("前払利息合計") = w前払利息合計        '2008/02/07 V182
                        wRs2("未払利息増合計") = w未払利息増合計    '2008/02/07 V182
                        wRs2("未払利息減合計") = w未払利息減合計    '2008/02/07 V182
                        wRs2("未払利息合計") = w未払利息合計        '2008/02/07 V182
    
    
                        For k = 1 To wcnt
                            wRs2("融資_" + CStr(Format(k, "00"))) = w融資(k)
                            wRs2("元金_" + CStr(Format(k, "00"))) = w元金(k)
                            wRs2("利息_" + CStr(Format(k, "00"))) = w利息(k)
                            wRs2("返済_" + CStr(Format(k, "00"))) = w返済(k)
                            wRs2("解約_" + CStr(Format(k, "00"))) = w解約(k)
                            wRs2("保証_" + CStr(Format(k, "00"))) = w保証(k)
                            wRs2("手数料_" + CStr(Format(k, "00"))) = w手数料(k)        ' 08/12/09 V189
                            
                            wRs2("初期手数料_" + CStr(Format(k, "00"))) = w初期手数料(k)
                            wRs2("元金手数料_" + CStr(Format(k, "00"))) = w元金手数料(k)
                            wRs2("利息手数料_" + CStr(Format(k, "00"))) = w利息手数料(k)
                            
                            
                            wRs2("残高_" + CStr(Format(k, "00"))) = w残高(k)
    
                            wRs2("前払利息増_" + CStr(Format(k, "00"))) = w前払利息増(k) '2008/02/07 V182
                            wRs2("前払利息減_" + CStr(Format(k, "00"))) = w前払利息減(k) '2008/02/07 V182
                            wRs2("前払利息_" + CStr(Format(k, "00"))) = w前払利息(k)     '2008/02/07 V182
                            wRs2("未払利息増_" + CStr(Format(k, "00"))) = w未払利息増(k) '2008/02/07 V182
                            wRs2("未払利息減_" + CStr(Format(k, "00"))) = w未払利息減(k) '2008/02/07 V182
                            wRs2("未払利息_" + CStr(Format(k, "00"))) = w未払利息(k)     '2008/02/07 V182
                            
                            wRs2("利率_" + CStr(Format(k, "00"))) = w利率(k)             '11/02/17
                        Next
    
                    wRs2.Update
                
                 Else
                        
                    '2010/06/18
                    p推移(wiCnt).借入番号 = w借入番号
                    
                    p推移(wiCnt).利子補給金フラグ = w利子補給金フラグ           '16/03/24 利子補給に伴う変更
                    
                    p推移(wiCnt).融資合計 = w融資合計
                    p推移(wiCnt).元金合計 = w元金合計
                    p推移(wiCnt).利息合計 = w利息合計
                    p推移(wiCnt).返済合計 = w返済合計
                    p推移(wiCnt).解約合計 = w解約合計
                    p推移(wiCnt).保証合計 = w保証合計
                    p推移(wiCnt).手数料合計 = w手数料合計
                    
                    p推移(wiCnt).初期手数料合計 = w初期手数料合計
                    p推移(wiCnt).元金手数料合計 = w元金手数料合計
                    p推移(wiCnt).利息手数料合計 = w利息手数料合計
                    
                    p推移(wiCnt).残高合計 = w残高合計

                    p推移(wiCnt).前払利息増合計 = w前払利息増合計
                    p推移(wiCnt).前払利息減合計 = w前払利息減合計
                    p推移(wiCnt).前払利息合計 = w前払利息合計
                    p推移(wiCnt).未払利息増合計 = w未払利息増合計
                    p推移(wiCnt).未払利息減合計 = w未払利息減合計
                    p推移(wiCnt).未払利息合計 = w未払利息合計
                    
                    p推移(wiCnt).損益利息額合計 = w損益利息額合計           '16/03/24
                    
                    
                    For k = 1 To wcnt
                        p推移(wiCnt).融資(k) = w融資(k)
                        p推移(wiCnt).元金(k) = w元金(k)
                        p推移(wiCnt).利息(k) = w利息(k)
                        p推移(wiCnt).返済(k) = w返済(k)
                        p推移(wiCnt).解約(k) = w解約(k)
                        p推移(wiCnt).保証(k) = w保証(k)
                        p推移(wiCnt).手数料(k) = w手数料(k)        ' 08/12/09 V189
                        
                        p推移(wiCnt).初期手数料(k) = w初期手数料(k)
                        p推移(wiCnt).元金手数料(k) = w元金手数料(k)
                        p推移(wiCnt).利息手数料(k) = w利息手数料(k)
                        
                        p推移(wiCnt).残高(k) = w残高(k)
                        
                        '***手入力長短振替額の算出ルーチン          ’16/01/25
                        If w残高(k) <> 0 Then                       '16/01/25
                            w対象年月 = DateAdd("yyyy", 1, w年月(k))       '16/01/25
                            w融資残高 = MBD010_借入金手入力残高(p借入計画マスタ, 1, w対象年月)      '16/01/25
                            w長短振替額(k) = w残高(k) - w融資残高                                   '16/01/25
                        Else                                                                        '16/01/25
                            w長短振替額(k) = 0                                                      '16/01/25
                        End If                                                                      '16/01/25
                        
                        p推移(wiCnt).長短振替額(k) = w長短振替額(k)                                 '16/01/25

                        p推移(wiCnt).前払利息増(k) = w前払利息増(k) '2008/02/07 V182
                        p推移(wiCnt).前払利息減(k) = w前払利息減(k) '2008/02/07 V182
                        p推移(wiCnt).前払利息(k) = w前払利息(k)     '2008/02/07 V182
                        p推移(wiCnt).未払利息増(k) = w未払利息増(k) '2008/02/07 V182
                        p推移(wiCnt).未払利息減(k) = w未払利息減(k) '2008/02/07 V182
                        p推移(wiCnt).未払利息(k) = w未払利息(k)     '2008/02/07 V182
                        
                        p推移(wiCnt).損益利息額(k) = w損益利息額(k) '16/03/24
                        
                        p推移(wiCnt).利率(k) = w利率(k)             '11/02/17
                        
                    Next
                    
                    p推移(wiCnt).長短振替額合計 = w長短振替額(wcnt)             '16/01/26

                    wiCnt = wiCnt + 1
                 End If
             End If
                
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing

    If FLG_Mdata = True Then
        wRs2.Close
        Set wRs = Nothing
    End If
'
    If FLG_Mdata <> True Then
        wstr2 = ""
        wstr2 = wstr2 + "Select * From DCDA010_借入残高推移表結果"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
        
            For wiCnt = 0 To 9999
            
                If p推移(wiCnt).借入番号 = "" Then
                    Exit For
                End If
                
                wRs2.AddNew
                
                  If p推移(wiCnt).利子補給金フラグ = 1 Then             '16/03/24 利子補給に伴う変更
                  'If p推移(wiCnt).借入金種別区分 = "04" Then             '16/03/24
                    '***利子補給の時の処理                                  '16/03/24
                    wRs2("借入番号") = p推移(wiCnt).借入番号
                    wRs2("融資合計") = 0                                    '16/03/24
                    wRs2("元金合計") = 0                                    '16/03/24
                    wRs2("利息合計") = 0                                    '16/03/24
                    wRs2("返済合計") = 0                                    '16/03/24
                    wRs2("解約合計") = 0                                    '16/03/24
                    wRs2("保証合計") = 0                                    '16/03,24
                    wRs2("手数料合計") = 0                                  '16/03/24
                    
                    wRs2("初期手数料合計") = 0                              '16/03/24
                    wRs2("元金手数料合計") = 0                              '16/03/24
                    wRs2("利息手数料合計") = 0                              '16/03/24
                    
                    wRs2("残高合計") = 0                                    '16/03/24
        
                    wRs2("前払利息増合計") = 0                              '16/03/24
                    wRs2("前払利息減合計") = 0                              '16/03/24
                    wRs2("前払利息合計") = 0                                '16/03/24
                    wRs2("未払利息増合計") = 0                              '16/03/24
                    wRs2("未払利息減合計") = 0                              '16/03/24
                    wRs2("未払利息合計") = p推移(wiCnt).未払利息合計
        
                    For k = 1 To wcnt
                        wRs2("融資_" + CStr(Format(k, "00"))) = 0           '16/03/24
                        wRs2("元金_" + CStr(Format(k, "00"))) = 0           '16/03/24
                        wRs2("利息_" + CStr(Format(k, "00"))) = 0           '16/03/24
                        wRs2("返済_" + CStr(Format(k, "00"))) = 0           '16/03/24
                        wRs2("解約_" + CStr(Format(k, "00"))) = 0           '16/03/24
                        wRs2("保証_" + CStr(Format(k, "00"))) = 0           '16/03/24
                        wRs2("手数料_" + CStr(Format(k, "00"))) = 0         '16/03/24
                        
                        wRs2("初期手数料_" + CStr(Format(k, "00"))) = 0     '16/03/24
                        wRs2("元金手数料_" + CStr(Format(k, "00"))) = 0     '16/03/24
                        wRs2("利息手数料_" + CStr(Format(k, "00"))) = 0     '16/03/24
                        
                        wRs2("残高_" + CStr(Format(k, "00"))) = 0           '16/93/24
        
                        wRs2("前払利息増_" + CStr(Format(k, "00"))) = 0     '16/03/24
                        wRs2("前払利息減_" + CStr(Format(k, "00"))) = 0     '16/03/24
                        wRs2("前払利息_" + CStr(Format(k, "00"))) = 0       '16/03/24
                        wRs2("未払利息増_" + CStr(Format(k, "00"))) = 0     '16/03/24
                        wRs2("未払利息減_" + CStr(Format(k, "00"))) = 0     '16/03/24
                        wRs2("未払利息_" + CStr(Format(k, "00"))) = 0       '16/03/24
                        
                        wRs2("利率_" + CStr(Format(k, "00"))) = p推移(wiCnt).利率(k)    '11/02/17
                    Next                                                    '16/03/24
                    
                  Else                                                                  '16/03/24
                    
                    wRs2("借入番号") = p推移(wiCnt).借入番号
                    wRs2("融資合計") = p推移(wiCnt).融資合計
                    wRs2("元金合計") = p推移(wiCnt).元金合計
                    wRs2("利息合計") = p推移(wiCnt).利息合計
                    wRs2("返済合計") = p推移(wiCnt).返済合計
                    wRs2("解約合計") = p推移(wiCnt).解約合計
                    wRs2("保証合計") = p推移(wiCnt).保証合計
                    wRs2("手数料合計") = p推移(wiCnt).手数料合計
                    
                    wRs2("初期手数料合計") = p推移(wiCnt).初期手数料合計
                    wRs2("元金手数料合計") = p推移(wiCnt).元金手数料合計
                    wRs2("利息手数料合計") = p推移(wiCnt).利息手数料合計
                    
                    wRs2("残高合計") = p推移(wiCnt).残高合計
        
                    wRs2("前払利息増合計") = p推移(wiCnt).前払利息増合計
                    wRs2("前払利息減合計") = p推移(wiCnt).前払利息減合計
                    wRs2("前払利息合計") = p推移(wiCnt).前払利息合計
                    wRs2("未払利息増合計") = p推移(wiCnt).未払利息増合計
                    wRs2("未払利息減合計") = p推移(wiCnt).未払利息減合計
                    wRs2("未払利息合計") = p推移(wiCnt).未払利息合計
        
                    For k = 1 To wcnt
                        wRs2("融資_" + CStr(Format(k, "00"))) = p推移(wiCnt).融資(k)
                        wRs2("元金_" + CStr(Format(k, "00"))) = p推移(wiCnt).元金(k)
                        wRs2("利息_" + CStr(Format(k, "00"))) = p推移(wiCnt).利息(k)
                        wRs2("返済_" + CStr(Format(k, "00"))) = p推移(wiCnt).返済(k)
                        wRs2("解約_" + CStr(Format(k, "00"))) = p推移(wiCnt).解約(k)
                        wRs2("保証_" + CStr(Format(k, "00"))) = p推移(wiCnt).保証(k)
                        wRs2("手数料_" + CStr(Format(k, "00"))) = p推移(wiCnt).手数料(k)
                        
                        wRs2("初期手数料_" + CStr(Format(k, "00"))) = p推移(wiCnt).初期手数料(k)
                        wRs2("元金手数料_" + CStr(Format(k, "00"))) = p推移(wiCnt).元金手数料(k)
                        wRs2("利息手数料_" + CStr(Format(k, "00"))) = p推移(wiCnt).利息手数料(k)
                        
                        wRs2("残高_" + CStr(Format(k, "00"))) = p推移(wiCnt).残高(k)
        
                        wRs2("前払利息増_" + CStr(Format(k, "00"))) = p推移(wiCnt).前払利息増(k)
                        wRs2("前払利息減_" + CStr(Format(k, "00"))) = p推移(wiCnt).前払利息減(k)
                        wRs2("前払利息_" + CStr(Format(k, "00"))) = p推移(wiCnt).前払利息(k)
                        wRs2("未払利息増_" + CStr(Format(k, "00"))) = p推移(wiCnt).未払利息増(k)
                        wRs2("未払利息減_" + CStr(Format(k, "00"))) = p推移(wiCnt).未払利息減(k)
                        wRs2("未払利息_" + CStr(Format(k, "00"))) = p推移(wiCnt).未払利息(k)
                        
                        wRs2("利率_" + CStr(Format(k, "00"))) = p推移(wiCnt).利率(k)    '11/02/17
                    Next                                                                '16/03/24
                        
                  End If                                                                '16/03/24
                  
                  
        
                wRs2.Update
        
            Next wiCnt
            
        wRs2.Close
        Set wRs2 = Nothing
        
        '***DCDA010_借入残高推移表結果2の作成       ’16/01/25
        wStr3 = ""
        wStr3 = wStr3 + "Select * From DCDA010_借入残高推移表結果2"
        Call AdoRecordsetOpen(GDb, wRs3, wStr3)                         '16/01/25
        
            For wiCnt = 0 To 9999                                       '16/01/25
            
                If p推移(wiCnt).借入番号 = "" Then                      '16/01/25
                    Exit For                                            '16/01/25
                End If                                                  '16/01/25
                
                wRs3.AddNew                                             '16/01/25
                    wRs3("借入番号") = p推移(wiCnt).借入番号            '16/01/25
                   
                    wRs3("長短振替額合計") = p推移(wiCnt).長短振替額合計            '16/01/25
                    wRs3("損益利息額合計") = p推移(wiCnt).損益利息額合計  '16/03/24
        
                    For k = 1 To wcnt                                   '16/01/25
                    
                        wRs3("長短振替額_" + CStr(Format(k, "00"))) = p推移(wiCnt).長短振替額(k)    '16/01/25
                        
                        wRs3("損益利息額_" + CStr(Format(k, "00"))) = p推移(wiCnt).損益利息額(k)    '16/03/24
                    Next                                                                            '16/01/25
                    
                    
                    If p推移(wiCnt).利子補給金フラグ = 1 Then       '16/03/24 利子補給に伴う変更
                    'If p推移(wiCnt).借入金種別区分 = "04" Then      '16/03/24
                        '***利子補給の時　損益利息額のみマイナス、その他はゼロ
                        wRs3("長短振替額合計") = 0                      '16/03/24
                        wRs3("損益利息額合計") = -p推移(wiCnt).損益利息額合計   '16/03/24
                        For k = 1 To wcnt                               '16/03/24
                            wRs3("長短振替額_" + CStr(Format(k, "00"))) = 0                             '16/03/24
                            wRs3("損益利息額_" + CStr(Format(k, "00"))) = -p推移(wiCnt).損益利息額(k)   '16/03/24
                        Next
                    End If                                                                          '16/03/24
                    
                wRs3.Update                                                                         '16/01/25
        
            Next wiCnt                                                                              '16/01/25
            
        wRs3.Close                                                                                  '16/01/25
        Set wRs3 = Nothing                                                                          '16/01/25
        
        
        
    End If
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MRB010_手入力借入残高表_ERR:
    pERR_MES = pPROGRAM_ID + "/ MRB010_手入力借入残高表() でエラー" + vbCrLf + vbCrLf + _
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
' MRB010_手入力未払前払  2011/02/11
'------------------------------------------------
Public Sub MRB010_手入力未払前払(pTbl As String, p借入番号 As String)
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String, wWhere As String
    
    Dim wiCnt As Integer
    Dim p推移(9999) As MRB010_借入金推移表
    Dim FLG_Mdata As Boolean
    
    Dim p借入計画マスタ As MAA910_借入金                                       '5/10/8 V129
    Dim w銀行マスタ As MAA030_銀行
    Dim w借入金 As MAA910_借入金
    
    Dim j As Integer, k As Integer
    Dim w千円単位 As Integer
    Dim w年 As Integer, w月 As Integer, w日 As Integer                         '5/8/18 V129
    Dim ww月 As Integer                                                        '5/9/2 V129
    Dim w間隔 As Integer, w回数 As Integer, wcnt As Integer
    
    '5/9/12 V129 解約の時 1回前の年月との差＞＝２　1回前の融資残更新
    Dim w月差 As Integer
    Dim w処理 As Integer                                                       '5/9/9 V129
    
    Dim w残高算出 As Double                                                  '5/8/18 V129
    Dim w融資残高 As Double                                                  '5/9/9 V129
    Dim w前月残合計 As Double, w前月残高(12) As Double                          ' 07/02/09 V180
    Dim w融資合計 As Double, w融資(12) As Double
    Dim w元金合計 As Double, w元金(12) As Double
    Dim w利息合計 As Double, w利息(12) As Double
    Dim w返済合計 As Double, w返済(12) As Double
    Dim w解約合計 As Double, w解約(12) As Double
    Dim w残高合計 As Double, w残高(12) As Double
    Dim w保証合計 As Double, w保証(12) As Double
    Dim w手数料合計 As Double, w手数料(12) As Double           ' 08/12/08 V189
    
    Dim w初期手数料合計 As Double, w初期手数料(12) As Double
    Dim w元金手数料合計 As Double, w元金手数料(12) As Double
    Dim w利息手数料合計 As Double, w利息手数料(12) As Double
    
    Dim w前払利息増合計 As Double, w前払利息増(12)          '2008/02/06 V182
    Dim w前払利息減合計 As Double, w前払利息減(12)          '2008/02/06 V182
    Dim w前払利息合計 As Double, w前払利息(12)              '2008/02/06 V182
    Dim w未払利息増合計 As Double, w未払利息増(12)          '2008/02/06 V182
    Dim w未払利息減合計 As Double, w未払利息減(12)          '2008/02/06 V182
    Dim w未払利息合計 As Double, w未払利息(12)              '2008/02/06 V182
    Dim w基準年月 As Date                                   '2008/02/06 V182
    Dim w終了年月 As Variant                                '2008/02/06 V182
    Dim w開始回目 As Integer                                '2008/02/07 V182
    Dim w終了回目 As Integer                                '2008/02/07 V182
    Dim w回目 As Integer                                    '2008/02/06 V182
    Dim w解約締切年月日 As Variant                          '2008/02/06 V182
    Dim w利息計算年月日 As Variant                          '10/02/04
    
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
    Dim w借入番号 As String, w借入計画番号 As String, w金融リストラ As String
    
    Dim w解約判定 As String
    Dim ws基本 As String                                                       '5/8/30 V129
    Dim wベンチャ As String, wベンチャcode As String
    
    Dim w借入貸付 As String                                                     ' 07/02/09 V180
    Dim w前月残 As Double                                                       ' 07/02/09 V180
    
    '5/10/9 V129 支店への貸付　リストラ番号
    Dim w期首借入年度 As String, w支店貸付 As String, w全社借入 As String
'    Dim wsTbl As String
'
    On Error GoTo MRB010_手入力未払前払_ERR
'
    ' -----------------------------------------
    '       借入金マスタより DCDA010_借入残高推移表結果　作成
    ' -----------------------------------------
    'w開始年 = GRpt.テキスト_01
    w開始年 = 22        '11/02/14
    w借入計画番号 = GRpt.借入
    w金融リストラ = GRpt.金融
    w千円単位 = GRpt.千円単位

    If w千円単位 = 1 Then
        w分母 = "1000"
    Else
        w分母 = "1"
    End If
    
    ReDim G利息未払前払テーブル(0)
    
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
    GRpt.推移 = "月次"                          '11/02/14
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
    FLG_Mdata = False '通常はデータ一括書込
    wiCnt = 0
'
    wstr = ""
    wstr = wstr + "Select * From " & pTbl
    wstr = wstr + " Where 借入番号 = '" & p借入番号 & "'"
    wstr = wstr + " And 手入力区分 = 1"                   ' 07/02/09 V180
    
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

    'If pTbl2 <> "" Then
    '    wstr = wstr + " UNION Select * From " & pTbl2
    '    wstr = wstr + " Where 手入力区分 = 1"
    '    wstr = wstr + " And ((借入計画番号='' And sm区分 = 0 And 取消フラグ = 0)"
    '    wstr = wstr + " Or (借入計画番号 = '" & w借入計画番号 & "' And 借入計画番号 <> ''"
    '    wstr = wstr + " And sm区分 = 0 And 取消フラグ = 0))"
    'End If
        
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.RecordCount >= 10000 Then
        FLG_Mdata = True
        
        wstr2 = ""
        wstr2 = wstr2 + "Select * From DCDA010_借入残高推移表結果"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
    End If
        
        Do Until wRs.eof
            p借入計画マスタ = MBD010_借入データセット(wRs)      '5/10/8 V129
         
            '** 借入金テーブル セット **
            'Call MBD010_借入金テーブル作成(w金融リストラ, p借入計画マスタ)
            
            w借入番号 = p借入計画マスタ.借入番号                '5/10/8 V129
            For j = 1 To 12
                w前月残高(j) = 0                                ' 07/02/09 V180
                w融資(j) = 0
                w元金(j) = 0
                w利息(j) = 0
                w返済(j) = 0
                w解約(j) = 0
                w残高(j) = 0
                w保証(j) = 0
                w手数料(j) = 0                      ' 08/12/08 V189
                
                w初期手数料(j) = 0
                w元金手数料(j) = 0
                w利息手数料(j) = 0
                
                w前払利息増(j) = 0                  '2008/02/06 V182
                w前払利息減(j) = 0                  '2008/02/06 V182
                w前払利息(j) = 0                    '2008/02/06 V182
                w未払利息増(j) = 0                  '2008/02/06 V182
                w未払利息減(j) = 0                  '2008/02/06 V182
                w未払利息(j) = 0                    '2008/02/06 V182
                
            Next
            
            w融資合計 = 0
            w元金合計 = 0
            w利息合計 = 0
            w返済合計 = 0
            w解約合計 = 0
            w残高合計 = 0
            w保証合計 = 0
            w手数料合計 = 0                         ' 08/12/08 V189
            
            w初期手数料合計 = 0
            w元金手数料合計 = 0
            w利息手数料合計 = 0
            
            w前払利息増合計 = 0                     '2008/02/06 V182
            w前払利息減合計 = 0                     '2008/02/06 V182
            w前払利息合計 = 0                       '2008/02/06 V182
            w未払利息増合計 = 0                     '2008/02/06 V182
            w未払利息減合計 = 0                     '2008/02/06 V182
            w未払利息合計 = 0                       '2008/02/06 V182
            
            For w回目 = 1 To 600                    '2008/02/06 V182
                p前払利息増(w回目) = 0              '2008/02/06 V182
                p前払利息減(w回目) = 0              '2008/02/06 V182
                p前払利息残(w回目) = 0              '2008/02/06 V182
                p未払利息増(w回目) = 0              '2008/02/06 V182
                p未払利息減(w回目) = 0              '2008/02/06 V182
                p未払利息残(w回目) = 0              '2008/02/06 V182
            Next                                    '2008/02/06 V182
            
            '
            
            ' =========================================
            '                 初期設定
            ' =========================================
            '期首前残設定
            
            w手入力区分 = p借入計画マスタ.手入力区分        '11/02/16
            
            '***
             Call MBD010_借入金入力明細Read(p借入計画マスタ)  ' 07/02/09 V180
            
            w対象年月 = DateAdd("m", 1, w年月(0))                   ' 07/02/09 V180
            w前月残 = MBD010_借入金手入力残高(p借入計画マスタ, 0, w対象年月) ' 07/02/09 V180
            w前月残高(1) = w前月残                                  ' 07/02/09 V180
            w対象年月 = MBA010_対象年月(CDate(p借入計画マスタ.実行日))  ' 07/02/09 V180
            
            w基準年月 = MBA010_対象年月(CDate(p借入計画マスタ.実行日))      '2008/02/06 V182
            
            
              
             For k = 1 To wcnt                                                               '5/10/8 V129
                 If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                     And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then '5/10/8 V129
                     w融資(k) = w融資(k) + p借入計画マスタ.融資金額                             '5/10/8 V129
                     Exit For                                                                '5/10/9 V129
                 End If                                                                      '5/10/8 V129
             Next                                                                            '5/10/8 V129
                
             '***手打ち入力　解約年月日SET　利息前払            ’10/01/13
             '   w解約締切年月日 = MBA010_締日年月日(CDate(w解約実行日))     '2008/02/06 V182
               
             w解約実行日 = Null                                 '10/01/13
             w解約締切年月日 = Null                             '10/01/13
             
             For j = 1 To UBound(G借入金入力)                    '10/01/13
                If p借入計画マスタ.利息区分 = "1" _
                        And Format(G借入金入力(j).借入返済年月日, "yyyy/mm/dd") _
                            = Format(p借入計画マスタ.最終返済実行日, "yyyy/mm/dd") _
                        And G借入金入力(j).利息額 < 0 Then              '10/01/13
                    w解約実行日 = p借入計画マスタ.最終返済実行日        '10/01/13
                    w解約締切年月日 = MBA010_締日年月日(CDate(w解約実行日))     '10/01/13
                    Exit For                                            '10/01/13
                End If                                                  '10/01/13
             Next                                                       '10/01/13
                    
                
             w前払利息残 = 0                                        '11/02/11
             w未払利息残 = 0                                        '11/02/11
                
                
             For j = 1 To UBound(G借入金入力)                        ' 07/02/09 V180
                w対象年月 = MBA010_対象年月(CDate(G借入金入力(j).借入返済年月日))
                For k = 1 To wcnt
                    If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                       And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then
                        w元金(k) = w元金(k) + G借入金入力(j).元金
                        w利息(k) = w利息(k) + G借入金入力(j).利息額
                        w返済(k) = w返済(k) + G借入金入力(j).返済金額
                        Exit For
                    End If
                Next
             
             '*****************************************************************
             '    利息前払　利息未払　メインルーチン  2008/02/06 V182
             '*****************************************************************
               
               'If G借入金入力 (j).利息額 <> 0 And _          2009/12/25
               '   ((Format(G借入金入力 (j).利息計算年月日, "yyyymmdd") _
               '    <> Format(p借入計画マスタ.最終返済実行日, "yyyymmdd")) Or _
               '    (Format(G借入金入力 (j).利息計算年月日, "yyyymmdd") _
               '    = Format(p借入計画マスタ.最終返済実行日, "yyyymmdd") And _
               '    Format(G借入金入力 (1).利息計算年月日, "yyyymmdd") _
               '    <> Format(p借入計画マスタ.実行日, "yyyymmdd"))) Then
                   
                
                'If Format(G借入金入力 (1).利息計算年月日, "yyyymmdd") _
                '     = Format(p借入計画マスタ.実行日, "yyyymmdd") And _
                '     G借入金入力 (1).利息額 <> 0 Then    '2009/12/15 V182
               If G借入金入力(j).利息額 <> 0 Or G借入金入力(j).仮計上利息額 Then                     ' 09/12/25
                'If p借入計画マスタ.利息区分 = "1" Then              ' 09/12/25
                '    w利息区分 = "1"                 '2008/02/07 V182
                '    If j = 1 Then                   '2008/02/07 V182
                '        w利息対象期間日数 = DateDiff("D", G借入金入力 (j).利息計算年月日, _
                '                                          G借入金入力 (j + 1).利息計算年月日) + 1  '2008/02/07 V182
                '    Else                            '2008/02/07 V182
                '        w利息対象期間日数 = DateDiff("D", G借入金入力 (j).利息計算年月日, _
                '                                          G借入金入力 (j + 1).利息計算年月日) '2008/02/07 V182
                '    End If                          '2008/02/07 V182
                    
                'Else                                '2008/02/07 V182
                '    w利息区分 = "2"                 '2008/02/07 V182
                '    If j = 1 Then                   '2008/07/07 V182
                '        w利息対象期間日数 = DateDiff("D", p借入計画マスタ.実行日, _
                '                                          G借入金入力 (j).利息計算年月日) + 1 '2008/02/07 V182
                '    Else                            '2008/02/07 V182
                '        w利息対象期間日数 = DateDiff("D", G借入金入力 (j - 1).利息計算年月日, _
                '                                          G借入金入力 (j).利息計算年月日) '2008/02/07 V182
                '    End If                          '2008/02/07 V182
                'End If                              '2008/02/07 V182
                
                    
                w対象年月 = MBA010_対象年月(CDate(G借入金入力(j).借入返済年月日))    '2008/02/07 V182
                w回目 = DateDiff("M", w基準年月, w対象年月) + 1     '2008/02/06 V182
                
                'w解約実行日 = Null                  '2008/02/07 V182
                'w解約締切年月日 = Null              '2008/02/07 V182
                 
                '日数調整、利息調整
                If G借入金入力(j).利息対象期間日数 = 0 Then          '10/01/08
                    w利息対象期間日数 = G借入金入力(j).日割日数      '10/01/08
                Else                                                '10/01/08
                    w利息対象期間日数 = G借入金入力(j).利息対象期間日数  '10/01/08
                End If                                              '10/01/08
                
                
                '***HEAD部　手打ちSET
                w利息未払前払.銀行番号 = p借入計画マスタ.銀行番号   '11/02/11
                w利息未払前払.借入番号 = p借入計画マスタ.借入番号   '11/02/11
                w利息未払前払.利息区分 = p借入計画マスタ.利息区分   '11/02/11
                w利息未払前払.利息計算日数区分 = p借入計画マスタ.利息計算日数区分   '11/02/11
                
                
                If p借入計画マスタ.利息区分 = "1" Then              '10/01/09 V182
                    If w回目 > 0 And w回目 <= 600 Then
                        p前払利息増(w回目) = p前払利息増(w回目) + G借入金入力(j).利息額  '2008/02/06 V182
                    End If
                    
                    If Format(G借入金入力(j).借入返済年月日, "yyyy/mm/dd") _
                        = Format(p借入計画マスタ.実行日, "yyyy/mm/dd") Then
                        If p借入計画マスタ.利息控除区分 = 1 Or p借入計画マスタ.利息控除区分 = 3 Then '10/02/04
                            w利息計算年月日 = DateAdd("d", 1, G借入金入力(j).利息計算年月日)                 '10/02/04
                        Else                                                                                '10/02/04
                            w利息計算年月日 = G借入金入力(j).利息計算年月日
                        End If                      '10/02/04
                        
                    Else                            '10/02/04
                        w利息計算年月日 = DateAdd("d", 1, G借入金入力(j).利息計算年月日)  '10/02/04
                    End If
                    
                    If Format(G借入金入力(j).借入返済年月日, "yyyymmdd") _
                       = Format(w解約実行日, "yyyymmdd") Then               '10/02/04
                        w利息計算年月日 = G借入金入力(j).利息計算年月日      '10/02/04
                    End If                                                  '10/02/04
                    
                    '***ITEM部　手打ちSET　前払増
                    w利息未払前払.返済年月日 = G借入金入力(j).借入返済年月日     '11/02/11
                    w利息未払前払.月毎NO = 1                                    '11/02/11
                    w利息未払前払.元金額 = G借入金入力(j).元金                   '11/02/11
                    w利息未払前払.融資残高 = G借入金入力(j).融資残高             '11/02/11
                    w利息未払前払.利息計算対象額 = G借入金入力(j).融資残高
                    w利息未払前払.利息額増 = G借入金入力(j).利息額               '11/02/11
                    w利息未払前払.利息額減 = 0                                  '11/02/11
                    w前払利息残 = w前払利息残 + G借入金入力(j).利息額            '11/02/11
                    w利息未払前払.利息残高 = w前払利息残                        '11/02/11
                    w利息未払前払.日割日数 = w利息対象期間日数                  '11/02/11
                    w利息未払前払.利率 = G借入金入力(j).利率                     '11/02/11
                    w利息未払前払.開始年月日 = w利息計算年月日                  '11/02/11
                    
                    If G借入金入力(j).日割日数 < 0 Then                          '11/02/11
                        w利息未払前払.終了年月日 = DateAdd("d", -G借入金入力(j).日割日数, w利息計算年月日)   '11/02/11
                    Else                                                        '11/02/11
                        w利息未払前払.終了年月日 = DateAdd("d", G借入金入力(j).日割日数 - 1, w利息計算年月日)  '11/02/11
                    End If                                                      '11/02/11
                    
                    w利息未払前払.利息期間対象日数 = w利息対象期間日数          '2014/08/26
                    w利息未払前払.利息期間対象額 = G借入金入力(j).利息額        '2014/08/26
                    w利息未払前払.利息調整F = 0                                     '2014/08/26
                    
                    Call MBD010_利息未払前払Write(w利息未払前払, p借入計画マスタ)   '2016/09/26
                    
                       
                    
                    Call MRB010_前払利息減計算明細(p借入計画マスタ, _
                                               p借入計画マスタ.実行日, _
                                               p借入計画マスタ.利息控除区分, _
                                               w利息計算年月日, _
                                               w利息対象期間日数, _
                                               G借入金入力(j).仮計上利息額 + G借入金入力(j).利息額, _
                                               w解約実行日, _
                                               w解約締切年月日, _
                                               w基準年月)     '2016/09/26
                    
                  
                Else                                                '2008/02/06 V182
                    If w回目 > 0 And w回目 <= 600 Then
                        p未払利息減(w回目) = p未払利息減(w回目) + G借入金入力(j).利息額  '2008/02/06 V182
                    End If
                    
                    If Format(G借入金入力(j).借入返済年月日, "yyyy/mm/dd") _
                        = Format(p借入計画マスタ.最終返済実行日, "yyyy/mm/dd") _
                        And (p借入計画マスタ.利息控除区分 = 2 Or p借入計画マスタ.利息控除区分 = 3) Then '10/02/04
                        w利息計算年月日 = DateAdd("D", -1, G借入金入力(j).利息計算年月日)                '10/02/04
                    Else                                                                                '10/02/04
                        w利息計算年月日 = G借入金入力(j).利息計算年月日                                  '10/02/04
                    End If
                    
                    '***ITEM部　手打ちSET　未払減
                    w利息未払前払.返済年月日 = G借入金入力(j).借入返済年月日     '11/02/11
                    w利息未払前払.月毎NO = 1                                    '11/02/11
                    w利息未払前払.元金額 = G借入金入力(j).元金                   '11/02/11
                    w利息未払前払.融資残高 = G借入金入力(j).融資残高             '11/02/11
                    w利息未払前払.利息計算対象額 = G借入金入力(j).融資残高 + G借入金入力(j).元金
                    w利息未払前払.利息額増 = 0                                  '11/02/11
                    w利息未払前払.利息額減 = G借入金入力(j).利息額               '11/02/11
                    w未払利息残 = w未払利息残 - G借入金入力(j).利息額            '11/02/11
                    w利息未払前払.利息残高 = w未払利息残                        '11/02/11
                    w利息未払前払.日割日数 = w利息対象期間日数                  '11/02/11
                    w利息未払前払.利率 = G借入金入力(j).利率                     '11/02/11
                    w利息未払前払.開始年月日 = DateAdd("d", 1 - w利息対象期間日数, w利息計算年月日) '11/02/11
                    w利息未払前払.終了年月日 = w利息計算年月日                  '11/02/11
                    
                    
                    w利息未払前払.利息期間対象日数 = w利息対象期間日数          '2014/08/26
                    w利息未払前払.利息期間対象額 = G借入金入力(j).利息額        '2014/08/26
                    w利息未払前払.利息調整F = 0                                 '2014/08/26
                    
                    Call MBD010_利息未払前払Write(w利息未払前払, p借入計画マスタ)   '2016/09/26
                    
                    
                    
                    Call MRB010_未払利息増計算明細(p借入計画マスタ, _
                                               p借入計画マスタ.実行日, _
                                               w利息計算年月日, _
                                               w利息対象期間日数, _
                                               G借入金入力(j).仮計上利息額 + G借入金入力(j).利息額, _
                                               w解約実行日, _
                                               w解約締切年月日, _
                                               w基準年月)     '2016/9/26
                    
                End If                                              '2008/02/06 V182
                
               End If                                           '09/12/26
               
               
             Next
             
             '*** DCDA030_利息未払前払明細　の作成
             Call MBD010_利息未払前払明細作成
             
             
             '***終了年月算出
             If Not IsNull(w解約実行日) Then                    '2008/02/07 V182
                w終了年月 = w解約実行日                         '2008/02/07 V182
             Else                                               '2008/02/07 V182
                w終了年月 = p借入計画マスタ.最終返済実行日      '2008/02/07 V182
             End If                                             '2008/02/07 V182
             
             w終了年月 = MBA010_対象年月(CDate(w終了年月))      '2008/02/07 V182
             
             w開始回目 = DateDiff("M", w基準年月, w基準年月) + 1 '2008/02/07 V182
             w終了回目 = DateDiff("M", w基準年月, w終了年月) + 1 '2008/02/07 V182
             
             '**************************************************************
             '     前払利息　未払利息　集計セット
             '**************************************************************
             For j = w開始回目 To w終了回目                     '2008/02/07 V182
                w対象年月 = DateAdd("M", j - 1, w基準年月)      '2008/02/07 V182
                If j = 1 Then                                   '2008/02/07 V182
                    p前払利息残(j) = p前払利息増(j) - p前払利息減(j) '2008/02/07V182
                    p未払利息残(j) = p未払利息増(j) - p未払利息減(j) '2008/02/07V182
                Else                                            '2008/02/07 V182
                    p前払利息残(j) = p前払利息残(j - 1) + p前払利息増(j) - p前払利息減(j) '2008/02/07 V182
                    p未払利息残(j) = p未払利息残(j - 1) + p未払利息増(j) - p未払利息減(j) '2008/02/07V182
                End If                                          '2008/02/07 V182
                
                For k = 1 To wcnt                               '2008/02/07 V182
                    If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                        And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then '2008/02/07 V182
                            
                        w前払利息増(k) = w前払利息増(k) + p前払利息増(j) '2008/02/07 V182
                        w前払利息減(k) = w前払利息減(k) + p前払利息減(j) '2008/02/07 V182
                        w未払利息増(k) = w未払利息増(k) + p未払利息増(j) '2008/02/07 V182
                        w未払利息減(k) = w未払利息減(k) + p未払利息減(j) '2008/02/07 V182
                        
                        If Format(w対象年月, "yyyymmdd") = Format(w年月(k), "yyyymmdd") Then '2008/02/07 V182
                            w前払利息(k) = p前払利息残(j)       '2008/02/07 V182
                            w未払利息(k) = p未払利息残(j)       '2008/02/07 V182
                        End If                                  '2008/02/07 V182
                    End If                                      '2008/02/07 V182
                Next                                            '2008/02/07 V182
             Next                                               '2008/02/07 V182
             
             
             
             '残高算出
             For k = 1 To wcnt
                If k > 1 Then
                    w前月残高(k) = w残高(k - 1)
                End If
                
                w残高(k) = w前月残高(k) + w融資(k) - w元金(k)
             Next
             
          

               
             For k = 1 To wcnt
                If w解約(k) <> 0 Then               '11/06/17 V200
                    w元金(k) = w元金(k) + w解約(k)  '11/06/17 V200
                    w返済(k) = w返済(k) + w解約(k)  '11/06/17 V200
                End If                              '11/06/17 V200
                
                w融資合計 = w融資合計 + w融資(k)
                w元金合計 = w元金合計 + w元金(k)
                w利息合計 = w利息合計 + w利息(k)
                w返済合計 = w返済合計 + w返済(k)
                w解約合計 = w解約合計 + w解約(k)
                w保証合計 = w保証合計 + w保証(k)
                w手数料合計 = w手数料合計 + w手数料(k)                  ' 08/12/08 V189
                
                w前払利息増合計 = w前払利息増合計 + w前払利息増(k)      '2008/02/07 V182
                w前払利息減合計 = w前払利息減合計 + w前払利息減(k)      '2008/02/07 V182
                w未払利息増合計 = w未払利息増合計 + w未払利息増(k)      '2008/02/07 V182
                w未払利息減合計 = w未払利息減合計 + w未払利息減(k)      '2008/02/07 V182
                
             Next
             w残高合計 = w残高(wcnt)
             
             w前払利息合計 = w前払利息(wcnt)                            '2008/02/07 V182
             w未払利息合計 = w未払利息(wcnt)                            '2008/02/07 V182
             
                
             If w融資合計 = 0 And w元金合計 = 0 And w利息合計 = 0 And _
                w返済合計 = 0 And w解約合計 = 0 And w保証合計 = 0 And _
                w前払利息増合計 = 0 And w前払利息減合計 = 0 And _
                w未払利息増合計 = 0 And w未払利息減合計 = 0 And _
                w残高合計 = 0 Then
             Else
                 If FLG_Mdata = True Then
                    wRs2.AddNew
                        wRs2("借入番号") = w借入番号
                        wRs2("融資合計") = w融資合計
                        wRs2("元金合計") = w元金合計
                        wRs2("利息合計") = w利息合計
                        wRs2("返済合計") = w返済合計
                        wRs2("解約合計") = w解約合計
                        wRs2("保証合計") = w保証合計
                        wRs2("手数料合計") = w手数料合計            ' 08/12/09 V189
                        
                        wRs2("初期手数料合計") = w初期手数料合計
                        wRs2("元金手数料合計") = w元金手数料合計
                        wRs2("利息手数料合計") = w利息手数料合計
                        
                        wRs2("残高合計") = w残高合計
    
                        wRs2("前払利息増合計") = w前払利息増合計    '2008/02/07 V182
                        wRs2("前払利息減合計") = w前払利息減合計    '2008/02/07 V182
                        wRs2("前払利息合計") = w前払利息合計        '2008/02/07 V182
                        wRs2("未払利息増合計") = w未払利息増合計    '2008/02/07 V182
                        wRs2("未払利息減合計") = w未払利息減合計    '2008/02/07 V182
                        wRs2("未払利息合計") = w未払利息合計        '2008/02/07 V182
    
    
                        For k = 1 To wcnt
                            wRs2("融資_" + CStr(Format(k, "00"))) = w融資(k)
                            wRs2("元金_" + CStr(Format(k, "00"))) = w元金(k)
                            wRs2("利息_" + CStr(Format(k, "00"))) = w利息(k)
                            wRs2("返済_" + CStr(Format(k, "00"))) = w返済(k)
                            wRs2("解約_" + CStr(Format(k, "00"))) = w解約(k)
                            wRs2("保証_" + CStr(Format(k, "00"))) = w保証(k)
                            wRs2("手数料_" + CStr(Format(k, "00"))) = w手数料(k)        ' 08/12/09 V189
                            
                            wRs2("初期手数料_" + CStr(Format(k, "00"))) = w初期手数料(k)
                            wRs2("元金手数料_" + CStr(Format(k, "00"))) = w元金手数料(k)
                            wRs2("利息手数料_" + CStr(Format(k, "00"))) = w利息手数料(k)
                            
                            wRs2("残高_" + CStr(Format(k, "00"))) = w残高(k)
    
                            wRs2("前払利息増_" + CStr(Format(k, "00"))) = w前払利息増(k) '2008/02/07 V182
                            wRs2("前払利息減_" + CStr(Format(k, "00"))) = w前払利息減(k) '2008/02/07 V182
                            wRs2("前払利息_" + CStr(Format(k, "00"))) = w前払利息(k)     '2008/02/07 V182
                            wRs2("未払利息増_" + CStr(Format(k, "00"))) = w未払利息増(k) '2008/02/07 V182
                            wRs2("未払利息減_" + CStr(Format(k, "00"))) = w未払利息減(k) '2008/02/07 V182
                            wRs2("未払利息_" + CStr(Format(k, "00"))) = w未払利息(k)     '2008/02/07 V182
                        Next
    
                    wRs2.Update
                
                 Else
                        
                    '2010/06/18
                    p推移(wiCnt).借入番号 = w借入番号
                    p推移(wiCnt).融資合計 = w融資合計
                    p推移(wiCnt).元金合計 = w元金合計
                    p推移(wiCnt).利息合計 = w利息合計
                    p推移(wiCnt).返済合計 = w返済合計
                    p推移(wiCnt).解約合計 = w解約合計
                    p推移(wiCnt).保証合計 = w保証合計
                    p推移(wiCnt).手数料合計 = w手数料合計
                    
                    p推移(wiCnt).初期手数料合計 = w初期手数料合計
                    p推移(wiCnt).元金手数料合計 = w元金手数料合計
                    p推移(wiCnt).利息手数料合計 = w利息手数料合計
                    
                    p推移(wiCnt).残高合計 = w残高合計

                    p推移(wiCnt).前払利息増合計 = w前払利息増合計
                    p推移(wiCnt).前払利息減合計 = w前払利息減合計
                    p推移(wiCnt).前払利息合計 = w前払利息合計
                    p推移(wiCnt).未払利息増合計 = w未払利息増合計
                    p推移(wiCnt).未払利息減合計 = w未払利息減合計
                    p推移(wiCnt).未払利息合計 = w未払利息合計

                    For k = 1 To wcnt
                        p推移(wiCnt).融資(k) = w融資(k)
                        p推移(wiCnt).元金(k) = w元金(k)
                        p推移(wiCnt).利息(k) = w利息(k)
                        p推移(wiCnt).返済(k) = w返済(k)
                        p推移(wiCnt).解約(k) = w解約(k)
                        p推移(wiCnt).保証(k) = w保証(k)
                        p推移(wiCnt).手数料(k) = w手数料(k)        ' 08/12/09 V189
                        
                        p推移(wiCnt).初期手数料(k) = w初期手数料(k)
                        p推移(wiCnt).元金手数料(k) = w元金手数料(k)
                        p推移(wiCnt).利息手数料(k) = w利息手数料(k)
                        
                        p推移(wiCnt).残高(k) = w残高(k)

                        p推移(wiCnt).前払利息増(k) = w前払利息増(k) '2008/02/07 V182
                        p推移(wiCnt).前払利息減(k) = w前払利息減(k) '2008/02/07 V182
                        p推移(wiCnt).前払利息(k) = w前払利息(k)     '2008/02/07 V182
                        p推移(wiCnt).未払利息増(k) = w未払利息増(k) '2008/02/07 V182
                        p推移(wiCnt).未払利息減(k) = w未払利息減(k) '2008/02/07 V182
                        p推移(wiCnt).未払利息(k) = w未払利息(k)     '2008/02/07 V182
                    Next

                    wiCnt = wiCnt + 1
                 End If
             End If
                
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing

    If FLG_Mdata = True Then
        wRs2.Close
        Set wRs = Nothing
    End If
'
    If FLG_Mdata <> True Then
        wstr2 = ""
        wstr2 = wstr2 + "Select * From DCDA010_借入残高推移表結果"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
        
            For wiCnt = 0 To 9999
            
                If p推移(wiCnt).借入番号 = "" Then
                    Exit For
                End If
                
                wRs2.AddNew
                    wRs2("借入番号") = p推移(wiCnt).借入番号
                    wRs2("融資合計") = p推移(wiCnt).融資合計
                    wRs2("元金合計") = p推移(wiCnt).元金合計
                    wRs2("利息合計") = p推移(wiCnt).利息合計
                    wRs2("返済合計") = p推移(wiCnt).返済合計
                    wRs2("解約合計") = p推移(wiCnt).解約合計
                    wRs2("保証合計") = p推移(wiCnt).保証合計
                    wRs2("手数料合計") = p推移(wiCnt).手数料合計
                    
                    wRs2("初期手数料合計") = p推移(wiCnt).初期手数料合計
                    wRs2("元金手数料合計") = p推移(wiCnt).元金手数料合計
                    wRs2("利息手数料合計") = p推移(wiCnt).利息手数料合計
                    
                    wRs2("残高合計") = p推移(wiCnt).残高合計
        
                    wRs2("前払利息増合計") = p推移(wiCnt).前払利息増合計
                    wRs2("前払利息減合計") = p推移(wiCnt).前払利息減合計
                    wRs2("前払利息合計") = p推移(wiCnt).前払利息合計
                    wRs2("未払利息増合計") = p推移(wiCnt).未払利息増合計
                    wRs2("未払利息減合計") = p推移(wiCnt).未払利息減合計
                    wRs2("未払利息合計") = p推移(wiCnt).未払利息合計
        
                    For k = 1 To wcnt
                        wRs2("融資_" + CStr(Format(k, "00"))) = p推移(wiCnt).融資(k)
                        wRs2("元金_" + CStr(Format(k, "00"))) = p推移(wiCnt).元金(k)
                        wRs2("利息_" + CStr(Format(k, "00"))) = p推移(wiCnt).利息(k)
                        wRs2("返済_" + CStr(Format(k, "00"))) = p推移(wiCnt).返済(k)
                        wRs2("解約_" + CStr(Format(k, "00"))) = p推移(wiCnt).解約(k)
                        wRs2("保証_" + CStr(Format(k, "00"))) = p推移(wiCnt).保証(k)
                        wRs2("手数料_" + CStr(Format(k, "00"))) = p推移(wiCnt).手数料(k)
                        
                        wRs2("初期手数料_" + CStr(Format(k, "00"))) = p推移(wiCnt).初期手数料(k)
                        wRs2("元金手数料_" + CStr(Format(k, "00"))) = p推移(wiCnt).元金手数料(k)
                        wRs2("利息手数料_" + CStr(Format(k, "00"))) = p推移(wiCnt).利息手数料(k)
                        
                        wRs2("残高_" + CStr(Format(k, "00"))) = p推移(wiCnt).残高(k)
        
                        wRs2("前払利息増_" + CStr(Format(k, "00"))) = p推移(wiCnt).前払利息増(k)
                        wRs2("前払利息減_" + CStr(Format(k, "00"))) = p推移(wiCnt).前払利息減(k)
                        wRs2("前払利息_" + CStr(Format(k, "00"))) = p推移(wiCnt).前払利息(k)
                        wRs2("未払利息増_" + CStr(Format(k, "00"))) = p推移(wiCnt).未払利息増(k)
                        wRs2("未払利息減_" + CStr(Format(k, "00"))) = p推移(wiCnt).未払利息減(k)
                        wRs2("未払利息_" + CStr(Format(k, "00"))) = p推移(wiCnt).未払利息(k)
                    Next
        
                wRs2.Update
        
            Next wiCnt
            
        wRs2.Close
        Set wRs2 = Nothing
    End If
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MRB010_手入力未払前払_ERR:
    pERR_MES = pPROGRAM_ID + "/ MRB010_手入力未払前払() でエラー" + vbCrLf + vbCrLf + _
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
' MRB010_標準入力借入残高表
'------------------------------------------------
Public Sub MRB010_標準入力借入残高表(pTbl As String)
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset, wRs3 As ADODB.Recordset    '16/01/25
    
    Dim wstr As String, wstr2 As String, wStr3 As String, wWhere As String          '16/01/25
    
    Dim wiCnt As Integer
    Dim p推移(9999) As MRB010_借入金推移表
    Dim FLG_Mdata As Boolean
    
    Dim p借入計画マスタ As MAA910_借入金                                       '5/10/8 V129
    Dim w銀行マスタ As MAA030_銀行
    Dim w借入金 As MAA910_借入金
    
    Dim j As Integer, k As Integer
    Dim w千円単位 As Integer
    Dim w年 As Integer, w月 As Integer, w日 As Integer                         '5/8/18 V129
    Dim ww月 As Integer                                                        '5/9/2 V129
    Dim w間隔 As Integer, w回数 As Integer, wcnt As Integer
    
    '5/9/12 V129 解約の時 1回前の年月との差＞＝２　1回前の融資残更新
    Dim w月差 As Integer
    Dim w処理 As Integer                                                       '5/9/9 V129
    
    Dim w残高年月 As Date                                                       '16/01/25
    
    Dim w残高算出 As Double                                                  '5/8/18 V129
    Dim w融資残高 As Double                                                  '5/9/9 V129
    Dim w前月残合計 As Double, w前月残高(12) As Double                          ' 07/02/09 V180
    Dim w融資合計 As Double, w融資(12) As Double
    Dim w元金合計 As Double, w元金(12) As Double
    Dim w利息合計 As Double, w利息(12) As Double
    Dim w返済合計 As Double, w返済(12) As Double
    Dim w解約合計 As Double, w解約(12) As Double
    Dim w残高合計 As Double, w残高(12) As Double
    Dim w保証合計 As Double, w保証(12) As Double
    Dim w手数料合計 As Double, w手数料(12) As Double
    
    Dim w前払利息増合計 As Double, w前払利息増(12)          '2008/02/06 V182
    Dim w前払利息減合計 As Double, w前払利息減(12)          '2008/02/06 V182
    Dim w前払利息合計 As Double, w前払利息(12)              '2008/02/06 V182
    Dim w未払利息増合計 As Double, w未払利息増(12)          '2008/02/06 V182
    Dim w未払利息減合計 As Double, w未払利息減(12)          '2008/02/06 V182
    Dim w未払利息合計 As Double, w未払利息(12)              '2008/02/06 V182
    
    Dim w長短振替額合計 As Double, w長短振替額(12)          '16/01/25
    
    Dim w損益利息額合計 As Double, w損益利息額(12)          '16/03/23
    
    Dim w利率(12) As Double                                 '11/02/17
    Dim w基準年月 As Date                                   '2008/02/06 V182
    Dim w終了年月 As Variant                                '2008/02/06 V182
    Dim w開始回目 As Integer                                '2008/02/07 V182
    Dim w終了回目 As Integer                                '2008/02/07 V182
    Dim w回目 As Integer                                    '2008/02/06 V182
    Dim w解約締切年月日 As Variant                          '2008/02/06 V182
    
    Dim w利息計算年月日 As Variant                          '2016/02/03
    
    Dim wd01 As Date
    Dim w実際年月 As Date
    Dim w開始年月日 As Date
    Dim w年月(12) As Date
    Dim w実際年月日OLD As Date                                                 '5/8/18 V129
    Dim w対象年月 As Date, w対象年月調整 As Date                               'V120
    Dim w対象年月OLD As Date, w対象年月NEW As Date                             '5/8/18 V129
    
    Dim w返済予定年月 As Date                               '10/11/6
    
    Dim w解約実行日 As Variant                                                 '5/10/8 V129
    Dim w管理年月1 As Variant, w管理年月2 As Variant, w管理年月3 As Variant    '5/9/8 V129
    Dim w実績年月1 As Variant, w実績年月2 As Variant, w実績年月3 As Variant    '5/9/8 V129
    Dim w実績年月日1 As Variant, w実績年月日2 As Variant, w実績年月日3 As Variant '5/9/8 V129
    Dim w集計年月 As Variant                                                   '5/10/8 V129
    
    Dim w分母 As String
    Dim w開始年 As String
    Dim w借入番号 As String, w借入計画番号 As String, w金融リストラ As String
    
    'Dim w借入金種別区分 As String           '16/03/24
    Dim w利子補給金フラグ As String           '16/03/24
    
    
    Dim w解約判定 As String
    Dim ws基本 As String                                                       '5/8/30 V129
    Dim wベンチャ As String, wベンチャcode As String
    
    Dim w借入貸付 As String                                                     ' 07/02/09 V180
    Dim w前月残 As Double                                                       ' 07/02/09 V180
    
    '5/10/9 V129 支店への貸付　リストラ番号
    Dim w期首借入年度 As String, w支店貸付 As String, w全社借入 As String
'    Dim wsTbl As String

    Dim w借入金管理区分 As String                                               '16/01/25

'
    On Error GoTo MRB010_標準入力借入残高表_ERR
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
    '** ワークファイル 削除 **
    wstr2 = ""
    wstr2 = wstr2 + "Delete * From DCDA010_借入残高推移表結果"
    GDb.Execute wstr2
    
    
    wStr3 = ""                                                      '16/01/26
    wStr3 = wStr3 + "Delete * From DCDA010_借入残高推移表結果2"     '16/01/26
    GDb.Execute wStr3                                               '16/01/26
    
    
    FLG_Mdata = False '通常はデータ一括書込
    wiCnt = 0
'
    wstr = ""
    wstr = wstr + "Select * From " & pTbl
    wstr = wstr + " Where 手入力区分 = 0"                   ' 07/02/09 V180
    
    '**ベンチャーの場合
    If (wベンチャcode < "a" Or wベンチャcode > "z") And w借入計画番号 <> "" Then
        wstr = wstr + " And ("                               ' 07/02/09 V180
        
        If w金融リストラ <> "" Then
            wstr = wstr + " (金融リストラ番号='" & w金融リストラ & "' And 取消フラグ=0) Or  "
        End If
        
        wstr = wstr + " (借入計画番号='" & w借入計画番号 & "' And 借入計画番号 <> ''"
        
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

    'If pTbl2 <> "" Then
    '    wstr = wstr + " UNION Select * From " & pTbl2
    '    wstr = wstr + " Where 手入力区分 = 1"
    '    wstr = wstr + " And ((借入計画番号='' And sm区分 = 0 And 取消フラグ = 0)"
    '    wstr = wstr + " Or (借入計画番号 = '" & w借入計画番号 & "' And 借入計画番号 <> ''"
    '    wstr = wstr + " And sm区分 = 0 And 取消フラグ = 0))"
    'End If
        
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.RecordCount >= 10000 Then
        FLG_Mdata = True
        
        wstr2 = ""
        wstr2 = wstr2 + "Select * From DCDA010_借入残高推移表結果"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
    End If
    
        Do Until wRs.eof
            p借入計画マスタ = MBD010_借入データセット(wRs)      '5/10/8 V129
         
            '** 借入金テーブル セット **
            
            Call MBD010_借入金テーブル作成(w金融リストラ, p借入計画マスタ)    ' 07/02/18 V180
            
            w借入番号 = p借入計画マスタ.借入番号                '5/10/8 V129
            
            'w借入金種別区分 = p借入計画マスタ.借入金種別区分    '16/03/24
            w利子補給金フラグ = p借入計画マスタ.利子補給金フラグ '16/03/24 利子補給に伴う変更
            
            For j = 1 To 12
                w前月残高(j) = 0                                ' 07/02/09 V180
                w融資(j) = 0
                w元金(j) = 0
                w利息(j) = 0
                w返済(j) = 0
                w解約(j) = 0
                w残高(j) = 0
                w保証(j) = 0
                w手数料(j) = 0                      ' 08/12/09 V189
                
                w長短振替額(j) = 0                  '16/01/25
                
                w損益利息額(j) = 0                  '16/03/24
                
                w前払利息増(j) = 0                  '2008/02/06 V182
                w前払利息減(j) = 0                  '2008/02/06 V182
                w前払利息(j) = 0                    '2008/02/06 V182
                w未払利息増(j) = 0                  '2008/02/06 V182
                w未払利息減(j) = 0                  '2008/02/06 V182
                w未払利息(j) = 0                    '2008/02/06 V182
                
                w利率(j) = 0                        '11/02/17
                
            Next
            
            w融資合計 = 0
            w元金合計 = 0
            w利息合計 = 0
            w返済合計 = 0
            w解約合計 = 0
            w残高合計 = 0
            w保証合計 = 0
            w手数料合計 = 0                         ' 08/12/09 V189
            
            w長短振替額合計 = 0                     '16/01/25
            
            w損益利息額合計 = 0                     '16/03/24
            
            w前払利息増合計 = 0                     '2008/02/06 V182
            w前払利息減合計 = 0                     '2008/02/06 V182
            w前払利息合計 = 0                       '2008/02/06 V182
            w未払利息増合計 = 0                     '2008/02/06 V182
            w未払利息減合計 = 0                     '2008/02/06 V182
            w未払利息合計 = 0                       '2008/02/06 V182
            
            For w回目 = 1 To 600                    '2008/02/06 V182
                p前払利息増(w回目) = 0              '2008/02/06 V182
                p前払利息減(w回目) = 0              '2008/02/06 V182
                p前払利息残(w回目) = 0              '2008/02/06 V182
                p未払利息増(w回目) = 0              '2008/02/06 V182
                p未払利息減(w回目) = 0              '2008/02/06 V182
                p未払利息残(w回目) = 0              '2008/02/06 V182
            Next                                    '2008/02/06 V182
            
            
            '
            
            ' =========================================
            '                 初期設定
            ' =========================================
            '期首前残設定
            
            '***
             'Call MBD010_借入金入力明細Read(p借入計画マスタ.借入番号, p借入計画マスタ.借入貸付)  ' 07/02/09 V180
            
            w対象年月 = DateAdd("m", 1, w年月(0))                   ' 07/02/09 V180
            'w前月残 = MBD010_借入金手入力残高(p借入計画マスタ, 0, w対象年月) ' 07/02/09 V180
            w前月残 = MBD010_借入金標準入力残高(p借入計画マスタ, w金融リストラ, 0, w対象年月, G基本情報.借入金管理区分) '07/02/26 V180
            w前月残高(1) = w前月残                                  ' 07/02/09 V180
            w対象年月 = MBA010_対象年月(CDate(p借入計画マスタ.実行日))  ' 07/02/09 V180
            
            w基準年月 = MBA010_対象年月(CDate(p借入計画マスタ.実行日))      '2008/02/06 V182
            
              
             For k = 1 To wcnt                                                               '5/10/8 V129
                 If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                     And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then '5/10/8 V129
                     w融資(k) = w融資(k) + p借入計画マスタ.融資金額                             '5/10/8 V129
                     Exit For                                                                '5/10/9 V129
                 End If                                                                      '5/10/8 V129
             Next             '5/10/8 V129
             
             ssw = 0            '2012/02/24
                 
             For j = 1 To UBound(G借入金テーブル)                   ' 07/02/18 V180
             
                Call MBA010_借入金年月算出(G借入金テーブル(j).返済予定年月, _
                    G借入金テーブル(j).実際年月日, p借入計画マスタ.支払日)  ' 07/02/12 V180
                    
                If G基本情報.借入金管理区分 = XMXA020_区分("借入金管理区分", "管理用") Then '07/02/18 V180
                    w対象年月 = MBA010_対象年月((CDate(G管理年月)))             ' 07/02/18 V180
                Else                                                            ' 07/02/180V180
                    w対象年月 = MBA010_対象年月((CDate(G実績年月)))             ' 07/02/18 V180
                End If                                                          ' 07/02/180V180
                
                '***初回返済年月日手打ちの時　G実績年月で算出　2010/11/06
                If G借入金テーブル(j).実際年月日 <> p借入計画マスタ.実行日 _
                    And G借入金テーブル(j).返済予定年月 = p借入計画マスタ.初回返済年月 Then '10/11/06
                    w返済予定年月 = p借入計画マスタ.初回返済年月
                    GDate2 = MXA030_翌営業年月日計算(w返済予定年月, p借入計画マスタ.支払日 _
                                                , p借入計画マスタ.営業日区分)       '10/11/06
                                                
                                                
                    If GDate2 <> p借入計画マスタ.初回返済実行日 Then   '10/11/06
                        w対象年月 = MBA010_対象年月((CDate(G実績年月)))             '10/11/06
                    End If                                                          '10/11/06
                End If
                
                
                '***最終返済年月日手打ちの時　G実績年月で算出　2010/11/06
                If G借入金テーブル(j).実際年月日 <> p借入計画マスタ.実行日 _
                    And G借入金テーブル(j).返済予定年月 = p借入計画マスタ.最終返済年月 Then '10/11/06
                    w返済予定年月 = p借入計画マスタ.最終返済年月
                    GDate2 = MXA030_翌営業年月日計算(w返済予定年月, p借入計画マスタ.支払日 _
                                                , p借入計画マスタ.営業日区分)       '10/11/06
                                                
                                                
                    If GDate2 <> p借入計画マスタ.最終返済実行日 Then   '10/11/06
                        w対象年月 = MBA010_対象年月((CDate(G実績年月)))             '10/11/06
                    End If                                                          '10/11/06
                End If
                
                        
                
                '解約算出
                If w金融リストラ > "" _
                        And w金融リストラ = p借入計画マスタ.金融リストラ番号 Then  ' 07/02/18 V180
                        w解約実行日 = p借入計画マスタ.金融解約実行日
                Else
                        w解約実行日 = p借入計画マスタ.解約実行日
                End If
                
                If Format(w解約実行日, "yyyymmdd") = _
                            Format(G借入金テーブル(j).実際年月日, "yyyymmdd") Then  ' 07/02/18 V180
                    w対象年月 = MBA010_対象年月((CDate(G実績年月)))             ' 07/02/18 V180
                End If                                                          ' 07/02/18 V180
                        
             
                'w対象年月 = MBA010_対象年月(CDate(G借入金入力(J).借入返済年月日))
                For k = 1 To wcnt
                    If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                       And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then
                        w元金(k) = w元金(k) + G借入金テーブル(j).元金額     ' 07/02/18 V180
                        w利息(k) = w利息(k) + G借入金テーブル(j).利息額     ' 07/02/18 V180
                        w返済(k) = w返済(k) + G借入金テーブル(j).返済金額   ' 07/02/18 V180
                        w保証(k) = w保証(k) + G借入金テーブル(j).保証料     ' 07/02/18 V180
                        w手数料(k) = w手数料(k) + G借入金テーブル(j).手数料 ' 08/12/09 V189
                        
                        If Format(w解約実行日, "yyyymmdd") = _
                            Format(G借入金テーブル(j).実際年月日, "yyyymmdd") Then  ' 07/02/18 V180
                            w解約(k) = w解約(k) + G借入金テーブル(j).融資残高       ' 07/02/18 V180
                        End If                                                  ' 07/02/18 V180
                        
                        w利率(k) = G借入金テーブル(j).利率                  '11/02/17
                        
                        '*直前利率設定　　1012/02/24
                        If ssw = 0 Then
                            If j = 1 And G借入金テーブル(j).利率 = 0 Then
                                w直前利率 = p借入計画マスタ.利率
                            End If
                            
                            If j >= 2 Then
                                If G借入金テーブル(j - 1).利率 = 0 Then
                                    w直前利率 = p借入計画マスタ.利率
                                Else
                                    w直前利率 = G借入金テーブル(j - 1).利率
                                End If
                            End If
                            
                            ssw = 1
                        End If
                        
                        Exit For
                    End If
                Next
              
             
             '*****************************************************************
             '    利息前払　利息未払　メインルーチン  2008/02/06 V182
             '*****************************************************************
               '2020/10/20 損益利息修正 ADD STR
               'If G借入金テーブル(j).利息額 <> 0 Then                 '2008/02/06 V182
               '2020/10/20 損益利息修正 ADD END
               '2020/10/20 損益利息修正 ADD STR
               If G借入金テーブル(j).利息額 <> 0 _
                  Or Format(G借入金テーブル(j).実際年月日, "yyyy/mm/dd") = Format(w解約実行日, "yyyy/mm/dd") Then
               '2020/10/20 損益利息修正 ADD END
               
                w対象年月 = MBA010_対象年月(CDate(G借入金テーブル(j).実際年月日))   '2008/02/06 V182
                w回目 = DateDiff("M", w基準年月, w対象年月) + 1     '2008/02/06 V182
                
                If Not IsNull(w解約実行日) Then
                    w解約締切年月日 = MBA010_締日年月日(CDate(w解約実行日))     '2008/02/06 V182
                Else
                    w解約締切年月日 = w解約実行日
                End If
                
                
                If p借入計画マスタ.利息区分 = "1" Then              '2008/02/06 V182
                    If w回目 > 0 And w回目 <= 600 Then
                        p前払利息増(w回目) = p前払利息増(w回目) + G借入金テーブル(j).利息額 '2008/02/06 V182
                    End If
                    
                    
                    
                    
                    
                    
                    '*** 標準入力借入残表固定日数より持ってきた　　　2016/2/3
                    
                    '*** 10/02/04
                    If Format(G借入金テーブル(j).実際年月日, "yyyymmdd") _
                       = Format(p借入計画マスタ.実行日, "yyyymmdd") Then
                        If p借入計画マスタ.利息控除区分 = 1 Or p借入計画マスタ.利息控除区分 = 3 Then
                            w利息計算年月日 = DateAdd("d", 1, G借入金テーブル(j).利息計算年月日) '10/02/04
                        Else                                                        '10/02/04
                            w利息計算年月日 = G借入金テーブル(j).利息計算年月日     '10/02/04
                        End If                                                      '10/02/04
                    Else                                                            '10/02/04
                        w利息計算年月日 = DateAdd("d", 1, G借入金テーブル(j).利息計算年月日)  '10/02/04
                    End If                                                          '10/02/04
                    
                    If Format(G借入金テーブル(j).実際年月日, "yyyymmdd") _
                       = Format(w解約実行日, "yyyymmdd") Then               '10/02/04
                        w利息計算年月日 = G借入金テーブル(j).利息計算年月日 '10/02/04
                    End If                                                  '10/02/04
                    
                    
                    
                    'w利息計算年月日 = G借入金テーブル(j).利息計算年月日
                    Call MRB010_前払利息減計算(p借入計画マスタ, _
                                               p借入計画マスタ.実行日, _
                                               w利息計算年月日, _
                                               G借入金テーブル(j).日割日数, _
                                               G借入金テーブル(j).利息額, _
                                               w解約実行日, _
                                               w解約締切年月日, _
                                               w基準年月)     '2016/09/27
                    
                                               
                Else                                                '2008/02/06 V182
                    If w回目 > 0 And w回目 <= 600 Then
                        p未払利息減(w回目) = p未払利息減(w回目) + G借入金テーブル(j).利息額 '2008/02/06 V182
                    End If
                    
                    
                    '*** 10/02/04
                    If (Format(G借入金テーブル(j).実際年月日, "yyyymmdd") _
                       = Format(p借入計画マスタ.最終返済実行日, "yyyymmdd") _
                       Or Format(G借入金テーブル(j).実際年月日, "yyyy/mm/dd") _
                          = Format(w解約実行日, "yyyy/mm/dd")) _
                       And _
                         (p借入計画マスタ.利息控除区分 = 2 Or p借入計画マスタ.利息控除区分 = 3) Then
                        w利息計算年月日 = DateAdd("d", -1, G借入金テーブル(j).利息計算年月日)
                    Else                    '10/02/04
                        w利息計算年月日 = G借入金テーブル(j).利息計算年月日     '10/02/04
                    End If
                    
                    Call MRB010_未払利息増計算(p借入計画マスタ, _
                                               p借入計画マスタ.実行日, _
                                               w利息計算年月日, _
                                               G借入金テーブル(j).日割日数, _
                                               G借入金テーブル(j).利息額, _
                                               w解約実行日, _
                                               w解約締切年月日, _
                                               w基準年月)     '2016/09/27
                    
                    
                    
                    
                    
                    
                    '****** 標準入力借入残表固定日数より持ってくる　以前の状況　　　2016/2/3
                    'Call MRB010_前払利息減計算(p借入計画マスタ.実行日, _
                    '                           G借入金テーブル(j).実際年月日, _
                    '                           G借入金テーブル(j).利息対象期間日数, _
                    '                           G借入金テーブル(j).利息額, _
                    '                           w解約実行日, _
                    '                           w解約締切年月日, _
                    '                           w基準年月)     '2008/02/06 V182
                    '
                                               
                'Else                                                '2008/02/06 V182
                '    If w回目 > 0 And w回目 <= 600 Then
                '        p未払利息減(w回目) = p未払利息減(w回目) + G借入金テーブル(j).利息額 '2008/02/06 V182
                '    End If
                    
                '    Call MRB010_未払利息増計算(p借入計画マスタ.実行日, _
                '                               G借入金テーブル(j).実際年月日, _
                '                               G借入金テーブル(j).利息対象期間日数, _
                '                               G借入金テーブル(j).利息額, _
                '                               w解約実行日, _
                '                               w解約締切年月日, _
                '                               w基準年月)     '2008/02/06 V182
                    
                End If                                              '2008/02/06 V182
                
               End If
               
             Next
             
             '***終了年月算出
             If Not IsNull(w解約実行日) Then                    '2008/02/07 V182
                w終了年月 = w解約実行日                         '2008/02/07 V182
             Else                                               '2008/02/07 V182
                w終了年月 = p借入計画マスタ.最終返済実行日      '2008/02/07 V182
             End If                                             '2008/02/07 V182
             
             w終了年月 = MBA010_対象年月(CDate(w終了年月))      '2008/02/07 V182
             
             w開始回目 = DateDiff("M", w基準年月, w基準年月) + 1 '2008/02/07 V182
             w終了回目 = DateDiff("M", w基準年月, w終了年月) + 1 '2008/02/07 V182
             
             '**************************************************************
             '     前払利息　未払利息　集計セット
             '**************************************************************
             For j = w開始回目 To w終了回目                     '2008/02/07 V182
                w対象年月 = DateAdd("M", j - 1, w基準年月)      '2008/02/07 V182
                If j = 1 Then                                   '2008/02/07 V182
                    p前払利息残(j) = p前払利息増(j) - p前払利息減(j) '2008/02/07V182
                    p未払利息残(j) = p未払利息増(j) - p未払利息減(j) '2008/02/07V182
                Else                                            '2008/02/07 V182
                    p前払利息残(j) = p前払利息残(j - 1) + p前払利息増(j) - p前払利息減(j) '2008/02/07 V182
                    p未払利息残(j) = p未払利息残(j - 1) + p未払利息増(j) - p未払利息減(j) '2008/02/07V182
                End If                                          '2008/02/07 V182
                
                For k = 1 To wcnt                               '2008/02/07 V182
                    If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                        And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then '2008/02/07 V182
                            
                        w前払利息増(k) = w前払利息増(k) + p前払利息増(j) '2008/02/07 V182
                        w前払利息減(k) = w前払利息減(k) + p前払利息減(j) '2008/02/07 V182
                        w未払利息増(k) = w未払利息増(k) + p未払利息増(j) '2008/02/07 V182
                        w未払利息減(k) = w未払利息減(k) + p未払利息減(j) '2008/02/07 V182
                        
                        If Format(w対象年月, "yyyymmdd") = Format(w年月(k), "yyyymmdd") Then '2008/02/07 V182
                            w前払利息(k) = p前払利息残(j)       '2008/02/07 V182
                            w未払利息(k) = p未払利息残(j)       '2008/02/07 V182
                        End If                                  '2008/02/07 V182
                    End If                                      '2008/02/07 V182
                Next                                            '2008/02/07 V182
             Next                                               '2008/02/07 V182
             
             For k = 1 To wcnt
                If k > 1 Then
                    w前月残高(k) = w残高(k - 1)
                End If
                
                w残高(k) = w前月残高(k) + w融資(k) - w元金(k) - w解約(k)    ' 07/02/18 V180
                
             Next
             
             For k = 1 To wcnt
                If w解約(k) <> 0 Then               '11/06/17 V200
                    w元金(k) = w元金(k) + w解約(k)  '11/06/17 V200
                    w返済(k) = w返済(k) + w解約(k)  '11/06/17 V200
                End If                              '11/06/17 V200
                
                w融資合計 = w融資合計 + w融資(k)
                w元金合計 = w元金合計 + w元金(k)
                w利息合計 = w利息合計 + w利息(k)
                w返済合計 = w返済合計 + w返済(k)
                w解約合計 = w解約合計 + w解約(k)
                w保証合計 = w保証合計 + w保証(k)
                w手数料合計 = w手数料合計 + w手数料(k)                  ' 08/12/09 V189
                w前払利息増合計 = w前払利息増合計 + w前払利息増(k)      '2008/02/07 V182
                w前払利息減合計 = w前払利息減合計 + w前払利息減(k)      '2008/02/07 V182
                w未払利息増合計 = w未払利息増合計 + w未払利息増(k)      '2008/02/07 V182
                w未払利息減合計 = w未払利息減合計 + w未払利息減(k)      '2008/02/07 V182
                
                If p借入計画マスタ.利息区分 = "1" Then          '16/03/24
                    w損益利息額(k) = w前払利息減(k)             '16/03/24
                Else                                            '16/03/24
                    w損益利息額(k) = w未払利息増(k)             '16/03/24
                End If                                          '16/03/24
                
                w損益利息額合計 = w損益利息額合計 + w損益利息額(k) '16/03/24
                
                
             Next
             
             w残高合計 = w残高(wcnt)
             w前払利息合計 = w前払利息(wcnt)                            '2008/02/07 V182
             w未払利息合計 = w未払利息(wcnt)                            '2008/02/07 V182
             
             
             '***利率= 0 の時　調整処理
             w現在利率 = w直前利率
             For k = 1 To wcnt
                If w融資(k) <> 0 Or w元金(k) <> 0 Or w利息(k) <> 0 _
                                 Or w解約(k) <> 0 Or w残高(k) <> 0 Then
                    If w利率(k) = 0 Then
                        w利率(k) = w現在利率
                    End If
                    
                    w現在利率 = w利率(k)
                End If
             Next
             
             
             
             '***決算用の利率調整 2012/02/24
             For k = 1 To wcnt
                If (w残高(k) <> 0 Or w元金(k) <> 0 Or w利息(k) <> 0) And w利率(k) = 0 Then
                    If k = 1 Then
                        w利率(k) = w利率(k + 1)
                    Else
                        w利率(k) = w利率(k - 1)
                    End If
                End If
             Next
             
             
             '期間利率
            '2016/10/17 損益利息一覧表 金利参照
            Dim w次回年月日 As Date
            Dim s As Integer
            
            
            If GRpt.帳票名 = "損益利息一覧表" Then
                For k = 1 To wcnt
                    w次回年月日 = MRB010_次回年月日算出(w年月(k))
                    
                    For j = 2 To 100
                        If IsNull(p借入計画マスタ.金利(j).金利変更x回目年月) Then
                            Exit For
                        End If
                    
                        If w次回年月日 >= p借入計画マスタ.金利(j).金利変更x回目年月 Then
                            w利率(k) = p借入計画マスタ.金利(j).金利x回目
                        Else
                            Exit For
                        End If
                    Next
                
                Next
            End If
                
             If w融資合計 = 0 And w元金合計 = 0 And w利息合計 = 0 And _
                w返済合計 = 0 And w解約合計 = 0 And w保証合計 = 0 And _
                w前払利息増合計 = 0 And w前払利息減合計 = 0 And _
                w未払利息増合計 = 0 And w未払利息減合計 = 0 And _
                w残高合計 = 0 Then
             Else
             
                 If FLG_Mdata = True Then
                    wRs2.AddNew
                        wRs2("借入番号") = w借入番号
                        wRs2("融資合計") = w融資合計
                        wRs2("元金合計") = w元金合計
                        wRs2("利息合計") = w利息合計
                        wRs2("返済合計") = w返済合計
                        wRs2("解約合計") = w解約合計
                        wRs2("保証合計") = w保証合計
                        wRs2("手数料合計") = w手数料合計            ' 08/12/09 V189
                        wRs2("残高合計") = w残高合計
    
                        wRs2("前払利息増合計") = w前払利息増合計    '2008/02/07 V182
                        wRs2("前払利息減合計") = w前払利息減合計    '2008/02/07 V182
                        wRs2("前払利息合計") = w前払利息合計        '2008/02/07 V182
                        wRs2("未払利息増合計") = w未払利息増合計    '2008/02/07 V182
                        wRs2("未払利息減合計") = w未払利息減合計    '2008/02/07 V182
                        wRs2("未払利息合計") = w未払利息合計        '2008/02/07 V182
    
    
                        For k = 1 To wcnt
                            wRs2("融資_" + CStr(Format(k, "00"))) = w融資(k)
                            wRs2("元金_" + CStr(Format(k, "00"))) = w元金(k)
                            wRs2("利息_" + CStr(Format(k, "00"))) = w利息(k)
                            wRs2("返済_" + CStr(Format(k, "00"))) = w返済(k)
                            wRs2("解約_" + CStr(Format(k, "00"))) = w解約(k)
                            wRs2("保証_" + CStr(Format(k, "00"))) = w保証(k)
                            wRs2("手数料_" + CStr(Format(k, "00"))) = w手数料(k)        ' 08/12/09 V189
                            wRs2("残高_" + CStr(Format(k, "00"))) = w残高(k)
    
                            wRs2("前払利息増_" + CStr(Format(k, "00"))) = w前払利息増(k) '2008/02/07 V182
                            wRs2("前払利息減_" + CStr(Format(k, "00"))) = w前払利息減(k) '2008/02/07 V182
                            wRs2("前払利息_" + CStr(Format(k, "00"))) = w前払利息(k)     '2008/02/07 V182
                            wRs2("未払利息増_" + CStr(Format(k, "00"))) = w未払利息増(k) '2008/02/07 V182
                            wRs2("未払利息減_" + CStr(Format(k, "00"))) = w未払利息減(k) '2008/02/07 V182
                            wRs2("未払利息_" + CStr(Format(k, "00"))) = w未払利息(k)     '2008/02/07 V182
                            
                            wRs2("利率_" + CStr(Format(k, "00"))) = w利率(k)             '11/02/17
                            
                        Next
                        
    
                    wRs2.Update
                
                 Else
                        
                    '2010/06/18
                    p推移(wiCnt).借入番号 = w借入番号
                    
                    'p推移(wiCnt).借入金種別区分 = w借入金種別区分       '16/03/24
                    p推移(wiCnt).利子補給金フラグ = w利子補給金フラグ    '16/03/24 利子補給に伴う変更
                    
                    p推移(wiCnt).融資合計 = w融資合計
                    p推移(wiCnt).元金合計 = w元金合計
                    p推移(wiCnt).利息合計 = w利息合計
                    p推移(wiCnt).返済合計 = w返済合計
                    p推移(wiCnt).解約合計 = w解約合計
                    p推移(wiCnt).保証合計 = w保証合計
                    p推移(wiCnt).手数料合計 = w手数料合計
                    p推移(wiCnt).残高合計 = w残高合計

                    p推移(wiCnt).前払利息増合計 = w前払利息増合計
                    p推移(wiCnt).前払利息減合計 = w前払利息減合計
                    p推移(wiCnt).前払利息合計 = w前払利息合計
                    p推移(wiCnt).未払利息増合計 = w未払利息増合計
                    p推移(wiCnt).未払利息減合計 = w未払利息減合計
                    p推移(wiCnt).未払利息合計 = w未払利息合計
                    
                    p推移(wiCnt).損益利息額合計 = w損益利息額合計     '16/03/24
                    
                    For k = 1 To wcnt
                        p推移(wiCnt).融資(k) = w融資(k)
                        p推移(wiCnt).元金(k) = w元金(k)
                        p推移(wiCnt).利息(k) = w利息(k)
                        p推移(wiCnt).返済(k) = w返済(k)
                        p推移(wiCnt).解約(k) = w解約(k)
                        p推移(wiCnt).保証(k) = w保証(k)
                        p推移(wiCnt).手数料(k) = w手数料(k)        ' 08/12/09 V189
                        p推移(wiCnt).残高(k) = w残高(k)
                        
                        '***標準入力長短振替額の算出ルーチン                        '16/03/10
                            If w残高(k) <> 0 Then                                       '16/01/25
                                If GRpt.借入金管理区分 = P8.FCDbl(XMXA020_区分("借入管理区分", "決算用")) Then  '16/03/10
                                    w借入金管理区分 = "1"                               '16/03/10
                                Else                                                    '16/03/10
                                    w借入金管理区分 = "0"                               '16/03/10
                                End If                                                  '16/03/10
                                
                                w借入金管理区分 = GRpt.借入金管理区分                    '16/03/10
                                
                                
                                w対象年月 = DateAdd("yyyy", 1, w年月(k))                   '16/01/25
                                w融資残高 = MBD010_借入金標準入力残高(p借入計画マスタ, w金融リストラ, 1, w対象年月, w借入金管理区分) '16/03/10
                                w長短振替額(k) = w残高(k) - w融資残高                   '16/01/25
                            Else                                                        '16/01/25
                                w長短振替額(k) = 0                                      '16/01/25
                            End If                                                      '16/01/25
                            
                            p推移(wiCnt).長短振替額(k) = w長短振替額(k)                 '16/01/25
                            
                            

                        p推移(wiCnt).前払利息増(k) = w前払利息増(k) '2008/02/07 V182
                        p推移(wiCnt).前払利息減(k) = w前払利息減(k) '2008/02/07 V182
                        p推移(wiCnt).前払利息(k) = w前払利息(k)     '2008/02/07 V182
                        p推移(wiCnt).未払利息増(k) = w未払利息増(k) '2008/02/07 V182
                        p推移(wiCnt).未払利息減(k) = w未払利息減(k) '2008/02/07 V182
                        p推移(wiCnt).未払利息(k) = w未払利息(k)     '2008/02/07 V182
                        
                        p推移(wiCnt).損益利息額(k) = w損益利息額(k) '16/03/24
                        
                        p推移(wiCnt).利率(k) = w利率(k)             '11/02/17
                        
                    Next
                    
                    p推移(wiCnt).長短振替額合計 = w長短振替額(wcnt)                 '16/01/26

                    wiCnt = wiCnt + 1
                 End If
             
             End If
                
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing

    If FLG_Mdata = True Then
        wRs2.Close
        Set wRs = Nothing
    End If
'
    If FLG_Mdata <> True Then
        wstr2 = ""
        wstr2 = wstr2 + "Select * From DCDA010_借入残高推移表結果"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
        
            For wiCnt = 0 To 9999
            
                If p推移(wiCnt).借入番号 = "" Then
                    Exit For
                End If
                
                wRs2.AddNew
                
                If p推移(wiCnt).利子補給金フラグ = 1 Then               '16/03/24 利子補給に伴う変更
                'If p推移(wiCnt).借入金種別区分 = "04" Then               '16/03/24
                    
                    '***利子補給の時の処理
                    wRs2("借入番号") = p推移(wiCnt).借入番号
                    wRs2("融資合計") = 0                            '16/03/24
                    wRs2("元金合計") = 0                            '16/03/24
                    wRs2("利息合計") = 0                            '16/03/24
                    wRs2("返済合計") = 0                            '16/03/24
                    wRs2("解約合計") = 0                            '16/03/24
                    wRs2("保証合計") = 0                            '16/03/24
                    wRs2("手数料合計") = 0                          '16/03/24
                    wRs2("残高合計") = 0                            '16/03/24
        
                    wRs2("前払利息増合計") = 0                      '16/03/24
                    wRs2("前払利息減合計") = 0                      '16/03/24
                    wRs2("前払利息合計") = 0                        '16/03/24
                    wRs2("未払利息増合計") = 0                      '16/03/24
                    wRs2("未払利息減合計") = 0                      '16/03/24
                    wRs2("未払利息合計") = 0                        '16/03/24
        
                    For k = 1 To wcnt
                        wRs2("融資_" + CStr(Format(k, "00"))) = 0   '16/03/24
                        wRs2("元金_" + CStr(Format(k, "00"))) = 0   '16/03/24
                        wRs2("利息_" + CStr(Format(k, "00"))) = 0   '16/03/24
                        wRs2("返済_" + CStr(Format(k, "00"))) = 0   '16/03/24
                        wRs2("解約_" + CStr(Format(k, "00"))) = 0   '16/03/24
                        wRs2("保証_" + CStr(Format(k, "00"))) = 0   '16/03/24
                        wRs2("手数料_" + CStr(Format(k, "00"))) = 0 '16/03/24
                        wRs2("残高_" + CStr(Format(k, "00"))) = 0   '16/03/24
        
                        wRs2("前払利息増_" + CStr(Format(k, "00"))) = 0 '16/03/24
                        wRs2("前払利息減_" + CStr(Format(k, "00"))) = 0 '16/03/24
                        wRs2("前払利息_" + CStr(Format(k, "00"))) = 0   '16/03/24
                        wRs2("未払利息増_" + CStr(Format(k, "00"))) = 0 '16/03/24
                        wRs2("未払利息減_" + CStr(Format(k, "00"))) = 0 '16/03/24
                        wRs2("未払利息_" + CStr(Format(k, "00"))) = 0   '16/03/24
                        
                        wRs2("利率_" + CStr(Format(k, "00"))) = p推移(wiCnt).利率(k)    '11/02/17
                   Next                                                                 '16/03/24
                   
                   
                Else                                                                    '16/03/24
                
                
                
                
                
                    wRs2("借入番号") = p推移(wiCnt).借入番号
                    wRs2("融資合計") = p推移(wiCnt).融資合計
                    wRs2("元金合計") = p推移(wiCnt).元金合計
                    wRs2("利息合計") = p推移(wiCnt).利息合計
                    wRs2("返済合計") = p推移(wiCnt).返済合計
                    wRs2("解約合計") = p推移(wiCnt).解約合計
                    wRs2("保証合計") = p推移(wiCnt).保証合計
                    wRs2("手数料合計") = p推移(wiCnt).手数料合計
                    wRs2("残高合計") = p推移(wiCnt).残高合計
        
                    wRs2("前払利息増合計") = p推移(wiCnt).前払利息増合計
                    wRs2("前払利息減合計") = p推移(wiCnt).前払利息減合計
                    wRs2("前払利息合計") = p推移(wiCnt).前払利息合計
                    wRs2("未払利息増合計") = p推移(wiCnt).未払利息増合計
                    wRs2("未払利息減合計") = p推移(wiCnt).未払利息減合計
                    wRs2("未払利息合計") = p推移(wiCnt).未払利息合計
        
                    For k = 1 To wcnt
                        wRs2("融資_" + CStr(Format(k, "00"))) = p推移(wiCnt).融資(k)
                        wRs2("元金_" + CStr(Format(k, "00"))) = p推移(wiCnt).元金(k)
                        wRs2("利息_" + CStr(Format(k, "00"))) = p推移(wiCnt).利息(k)
                        wRs2("返済_" + CStr(Format(k, "00"))) = p推移(wiCnt).返済(k)
                        wRs2("解約_" + CStr(Format(k, "00"))) = p推移(wiCnt).解約(k)
                        wRs2("保証_" + CStr(Format(k, "00"))) = p推移(wiCnt).保証(k)
                        wRs2("手数料_" + CStr(Format(k, "00"))) = p推移(wiCnt).手数料(k)
                        wRs2("残高_" + CStr(Format(k, "00"))) = p推移(wiCnt).残高(k)
        
                        wRs2("前払利息増_" + CStr(Format(k, "00"))) = p推移(wiCnt).前払利息増(k)
                        wRs2("前払利息減_" + CStr(Format(k, "00"))) = p推移(wiCnt).前払利息減(k)
                        wRs2("前払利息_" + CStr(Format(k, "00"))) = p推移(wiCnt).前払利息(k)
                        wRs2("未払利息増_" + CStr(Format(k, "00"))) = p推移(wiCnt).未払利息増(k)
                        wRs2("未払利息減_" + CStr(Format(k, "00"))) = p推移(wiCnt).未払利息減(k)
                        wRs2("未払利息_" + CStr(Format(k, "00"))) = p推移(wiCnt).未払利息(k)
                        
                        wRs2("利率_" + CStr(Format(k, "00"))) = p推移(wiCnt).利率(k)    '11/02/17
                        
                    Next                                                '16/03/24
                End If                                                  '16/03/24
                
                    
        
                wRs2.Update
        
            Next wiCnt
            
        wRs2.Close
        Set wRs2 = Nothing
        
        
        
        
        wStr3 = ""                                                  '16/01/25
        wStr3 = wStr3 + "Select * From DCDA010_借入残高推移表結果2" '16/01/25
        Call AdoRecordsetOpen(GDb, wRs3, wStr3)                     '16/01/25
        
            For wiCnt = 0 To 9999                                   '16/01/25
            
                If p推移(wiCnt).借入番号 = "" Then                  '16/01/25
                    Exit For                                        '16/01/25
                End If
                
                wRs3.AddNew                                         '16/01/25
                    wRs3("借入番号") = p推移(wiCnt).借入番号        '16/01/25
                    
                    wRs3("長短振替額合計") = p推移(wiCnt).長短振替額合計                '16/01/25
                    wRs3("損益利息額合計") = p推移(wiCnt).損益利息額合計                '16/03/24
                    
                    For k = 1 To wcnt                                                   '16/01/25
                        
                        wRs3("長短振替額_" + CStr(Format(k, "00"))) = p推移(wiCnt).長短振替額(k)    '16/01/25
                        wRs3("損益利息額_" + CStr(Format(k, "00"))) = p推移(wiCnt).損益利息額(k)    '16/03/24
                    Next                                            '16/01/25
                    
                    If p推移(wiCnt).利子補給金フラグ = 1 Then   '16/03/24 利子補給に伴う変更
                    'If p推移(wiCnt).借入金種別区分 = "04" Then   '16/03/24
                        '***利子補給の時
                        wRs3("長短振替額合計") = 0                  '16/01/25
                        wRs3("損益利息額合計") = -p推移(wiCnt).損益利息額合計                '16/03/24
                        
                        For k = 1 To wcnt                           '16/03/24
                            wRs3("長短振替額_" + CStr(Format(k, "00"))) = 0             '16/03/24
                            wRs3("損益利息額_" + CStr(Format(k, "00"))) = -p推移(wiCnt).損益利息額(k)    '16/03/24
                        Next                                        '16/03/24
                    End If                                          '16/03/23
        
                wRs3.Update                                         '16/01/25
        
            Next wiCnt                                              '16/01/25
            
        wRs3.Close                                                  '16/01/25
        Set wRs3 = Nothing                                          '16/01/25
        
        
        
        
    End If
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MRB010_標準入力借入残高表_ERR:
    pERR_MES = pPROGRAM_ID + "/ MRB010_標準入力借入残高表() でエラー" + vbCrLf + vbCrLf + _
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
' MRB010_標準入力借入残高表固定日数
'------------------------------------------------
Public Sub MRB010_標準入力借入残高表固定日数(pTbl As String)
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String, wWhere As String
    
    Dim wiCnt As Integer
    Dim p推移(9999) As MRB010_借入金推移表
    Dim FLG_Mdata As Boolean
    
    Dim p借入計画マスタ As MAA910_借入金                                       '5/10/8 V129
    Dim w銀行マスタ As MAA030_銀行
    Dim w借入金 As MAA910_借入金
    
    Dim j As Integer, k As Integer
    Dim w千円単位 As Integer
    Dim w年 As Integer, w月 As Integer, w日 As Integer                         '5/8/18 V129
    Dim ww月 As Integer                                                        '5/9/2 V129
    Dim w間隔 As Integer, w回数 As Integer, wcnt As Integer
    
    '5/9/12 V129 解約の時 1回前の年月との差＞＝２　1回前の融資残更新
    Dim w月差 As Integer
    Dim w処理 As Integer                                                       '5/9/9 V129
    
    Dim w残高算出 As Double                                                  '5/8/18 V129
    Dim w融資残高 As Double                                                  '5/9/9 V129
    Dim w前月残合計 As Double, w前月残高(12) As Double                          ' 07/02/09 V180
    Dim w融資合計 As Double, w融資(12) As Double
    Dim w元金合計 As Double, w元金(12) As Double
    Dim w利息合計 As Double, w利息(12) As Double
    Dim w返済合計 As Double, w返済(12) As Double
    Dim w解約合計 As Double, w解約(12) As Double
    Dim w残高合計 As Double, w残高(12) As Double
    Dim w保証合計 As Double, w保証(12) As Double
    Dim w手数料合計 As Double, w手数料(12) As Double        ' 08/12/09 V189
    
    Dim w初期手数料合計 As Double, w初期手数料(12) As Double
    Dim w元金手数料合計 As Double, w元金手数料(12) As Double
    Dim w利息手数料合計 As Double, w利息手数料(12) As Double
    
    
    Dim w前払利息増合計 As Double, w前払利息増(12)          '2008/02/06 V182
    Dim w前払利息減合計 As Double, w前払利息減(12)          '2008/02/06 V182
    Dim w前払利息合計 As Double, w前払利息(12)              '2008/02/06 V182
    Dim w未払利息増合計 As Double, w未払利息増(12)          '2008/02/06 V182
    Dim w未払利息減合計 As Double, w未払利息減(12)          '2008/02/06 V182
    Dim w未払利息合計 As Double, w未払利息(12)              '2008/02/06 V182
    Dim w利率(12) As Double                                 '11/02/17
    Dim w基準年月 As Date                                   '2008/02/06 V182
    Dim w終了年月 As Variant                                '2008/02/06 V182
    Dim w開始回目 As Integer                                '2008/02/07 V182
    Dim w終了回目 As Integer                                '2008/02/07 V182
    Dim w回目 As Integer                                    '2008/02/06 V182
    Dim w解約締切年月日 As Variant                          '2008/02/06 V182
    Dim w利息計算年月日 As Variant                          '10/02/04
        
    Dim wd01 As Date
    Dim w実際年月 As Date
    Dim w実際年月日 As Variant                              '08/03/05 V185
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
    Dim w借入番号 As String, w借入計画番号 As String, w金融リストラ As String
    
    Dim w解約判定 As String
    Dim ws基本 As String                                                       '5/8/30 V129
    Dim wベンチャ As String, wベンチャcode As String
    
    Dim w借入貸付 As String                                                     ' 07/02/09 V180
    Dim w前月残 As Double                                                       ' 07/02/09 V180
    
    Dim w手打年月日 As Date                                                     ' 08/12/23 V189
    Dim wsw As Integer                                                          ' 08/12/23 V189
    
    '5/10/9 V129 支店への貸付　リストラ番号
    Dim w期首借入年度 As String, w支店貸付 As String, w全社借入 As String
'    Dim wsTbl As String
'
    On Error GoTo MRB010_標準入力借入残高表固定日数_ERR
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
    
    
    '** ワークファイル 削除 **
    wstr2 = ""
    wstr2 = wstr2 + "Delete * From DCDA010_借入残高推移表結果"
    GDb.Execute wstr2
    
    
    FLG_Mdata = False '通常はデータ一括書込
    wiCnt = 0
'
    wstr = ""
    wstr = wstr + "Select * From " & pTbl
    wstr = wstr + " Where 手入力区分 = 0"                   ' 07/02/09 V180
    
    '**ベンチャーの場合
    If (wベンチャcode < "a" Or wベンチャcode > "z") And w借入計画番号 <> "" Then
        wstr = wstr + " And ("                               ' 07/02/09 V180
        
        If w金融リストラ <> "" Then
            wstr = wstr + " (金融リストラ番号='" & w金融リストラ & "' And 取消フラグ=0) Or  "
        End If
        
        wstr = wstr + " (借入計画番号='" & w借入計画番号 & "' And 借入計画番号 <> ''"
        
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
        wstr = wstr + " Or (借入計画番号='" & w借入計画番号 & "' And 借入計画番号 <> ''"
        
        If G基本情報.企業区分 = "連結本部" Then         '06/09/22 V170
            wstr = wstr + " And Format(実行日,'yyyymmdd') > '" & Format(Gコントロール.最終実績年月, "yyyymmdd") & "'"
        End If                                    '17/8/17 V128
        
        wstr = wstr + " And 取消フラグ=0)"
        
        If G基本情報.企業区分 = "連結本部" Or G基本情報.企業区分 = "連結子会社" Or G基本情報.企業区分 = "全社" Then
            wstr = wstr + " Or (借入計画番号='" & w借入計画番号 & "' And sm区分=1 And 取消フラグ=0)"
        End If
    
    End If
        wstr = wstr + ")"

    'If pTbl2 <> "" Then
    '    wstr = wstr + " UNION Select * From " & pTbl2
    '    wstr = wstr + " Where 手入力区分 = 1"
    '    wstr = wstr + " And ((借入計画番号='' And sm区分 = 0 And 取消フラグ = 0)"
    '    wstr = wstr + " Or (借入計画番号 = '" & w借入計画番号 & "' And 借入計画番号 <> ''"
    '    wstr = wstr + " And sm区分 = 0 And 取消フラグ = 0))"
    'End If
        
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.RecordCount >= 10000 Then
        FLG_Mdata = True
        
        wstr2 = ""
        wstr2 = wstr2 + "Select * From DCDA010_借入残高推移表結果"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
    End If
        
        Do Until wRs.eof
            p借入計画マスタ = MBD010_借入データセット(wRs)      '5/10/8 V129
         
            '** 借入金テーブル セット **
            
            Call MBD010_借入金テーブル作成(w金融リストラ, p借入計画マスタ)    ' 07/02/18 V180
            
            w借入番号 = p借入計画マスタ.借入番号                '5/10/8 V129
            For j = 1 To 12
                w前月残高(j) = 0                                ' 07/02/09 V180
                w融資(j) = 0
                w元金(j) = 0
                w利息(j) = 0
                w返済(j) = 0
                w解約(j) = 0
                w残高(j) = 0
                w保証(j) = 0
                w手数料(j) = 0                      ' 08/12/09 V189
                
                w初期手数料(j) = 0
                w元金手数料(j) = 0
                w利息手数料(j) = 0
                
                w前払利息増(j) = 0                  '2008/02/06 V182
                w前払利息減(j) = 0                  '2008/02/06 V182
                w前払利息(j) = 0                    '2008/02/06 V182
                w未払利息増(j) = 0                  '2008/02/06 V182
                w未払利息減(j) = 0                  '2008/02/06 V182
                w未払利息(j) = 0                    '2008/02/06 V182
                
                w利率(j) = 0                        '11/02/17
                
            Next
            
            w融資合計 = 0
            w元金合計 = 0
            w利息合計 = 0
            w返済合計 = 0
            w解約合計 = 0
            w残高合計 = 0
            w保証合計 = 0
            w手数料合計 = 0                         ' 08/12/09 V189
            
            w初期手数料合計 = 0
            w元金手数料合計 = 0
            w利息手数料合計 = 0
            
            w前払利息増合計 = 0                     '2008/02/06 V182
            w前払利息減合計 = 0                     '2008/02/06 V182
            w前払利息合計 = 0                       '2008/02/06 V182
            w未払利息増合計 = 0                     '2008/02/06 V182
            w未払利息減合計 = 0                     '2008/02/06 V182
            w未払利息合計 = 0                       '2008/02/06 V182
            
            For w回目 = 1 To 600                    '2008/02/06 V182
                p前払利息増(w回目) = 0              '2008/02/06 V182
                p前払利息減(w回目) = 0              '2008/02/06 V182
                p前払利息残(w回目) = 0              '2008/02/06 V182
                p未払利息増(w回目) = 0              '2008/02/06 V182
                p未払利息減(w回目) = 0              '2008/02/06 V182
                p未払利息残(w回目) = 0              '2008/02/06 V182
            Next                                    '2008/02/06 V182
            
            
            '
            
            ' =========================================
            '                 初期設定
            ' =========================================
            '期首前残設定
            
            '***
             'Call MBD010_借入金入力明細Read(p借入計画マスタ.借入番号, p借入計画マスタ.借入貸付)  ' 07/02/09 V180
            
            w対象年月 = DateAdd("m", 1, w年月(0))                   ' 07/02/09 V180
            'w前月残 = MBD010_借入金手入力残高(p借入計画マスタ, 0, w対象年月) ' 07/02/09 V180
            w前月残 = MBD010_借入金標準入力残高(p借入計画マスタ, w金融リストラ, 0, w対象年月, G基本情報.借入金管理区分) '07/02/26 V180
            w前月残高(1) = w前月残                                  ' 07/02/09 V180
            w対象年月 = MBA010_対象年月(CDate(p借入計画マスタ.実行日))  ' 07/02/09 V180
            
            w基準年月 = MBA010_対象年月(CDate(p借入計画マスタ.実行日))      '2008/02/06 V182
            
              
             For k = 1 To wcnt                                                               '5/10/8 V129
                 If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                     And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then '5/10/8 V129
                     w融資(k) = w融資(k) + p借入計画マスタ.融資金額                             '5/10/8 V129
                     Exit For                                                                '5/10/9 V129
                 End If                                                                      '5/10/8 V129
             Next                                                                            '5/10/8 V129
                 
             ssw = 0    '2012/04/24
                 
                 
             For j = 1 To UBound(G借入金テーブル)                   ' 07/02/18 V180
             
                Call MBA010_借入金年月算出(G借入金テーブル(j).返済予定年月, _
                    G借入金テーブル(j).実際年月日, p借入計画マスタ.支払日)  ' 07/02/12 V180
                    
                If G基本情報.借入金管理区分 = XMXA020_区分("借入金管理区分", "管理用") Then '07/02/18 V180
                    w対象年月 = MBA010_対象年月((CDate(G管理年月)))             ' 07/02/18 V180
                Else                                                            ' 07/02/180V180
                    w対象年月 = MBA010_対象年月((CDate(G実績年月)))             ' 07/02/18 V180
                End If                                                          ' 07/02/180V180
                
                '解約算出
                If w金融リストラ > "" _
                        And w金融リストラ = p借入計画マスタ.金融リストラ番号 Then  ' 07/02/18 V180
                        w解約実行日 = p借入計画マスタ.金融解約実行日
                Else
                        w解約実行日 = p借入計画マスタ.解約実行日
                End If
                
                If Format(w解約実行日, "yyyymmdd") = _
                            Format(G借入金テーブル(j).実際年月日, "yyyymmdd") Then  ' 07/02/18 V180
                    w対象年月 = MBA010_対象年月((CDate(G実績年月)))             ' 07/02/18 V180
                End If                                                          ' 07/02/18 V180
                        
             
                'w対象年月 = MBA010_対象年月(CDate(G借入金テーブル (J).借入返済年月日))
                For k = 1 To wcnt
                    If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                       And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then
                        w元金(k) = w元金(k) + G借入金テーブル(j).元金額     ' 07/02/18 V180
                        w利息(k) = w利息(k) + G借入金テーブル(j).利息額     ' 07/02/18 V180
                        w返済(k) = w返済(k) + G借入金テーブル(j).返済金額   ' 07/02/18 V180
                        w保証(k) = w保証(k) + G借入金テーブル(j).保証料     ' 07/02/18 V180
                        w手数料(k) = w手数料(k) + G借入金テーブル(j).手数料 ' 08/12009 V189
                        
                        If Format(w解約実行日, "yyyymmdd") = _
                            Format(G借入金テーブル(j).実際年月日, "yyyymmdd") Then  ' 07/02/18 V180
                            w解約(k) = w解約(k) + G借入金テーブル(j).融資残高       ' 07/02/18 V180
                        End If                                                  ' 07/02/18 V180
                        
                        w利率(k) = G借入金テーブル(j).利率                  '11/02/17
                        
                        
                         '*直前利率設定　　1012/02/24
                        If ssw = 0 Then
                            If j = 1 And G借入金テーブル(j).利率 = 0 Then
                                w直前利率 = p借入計画マスタ.利率
                            End If
                            
                            If j >= 2 Then
                                If G借入金テーブル(j - 1).利率 = 0 Then
                                    w直前利率 = p借入計画マスタ.利率
                                Else
                                    w直前利率 = G借入金テーブル(j - 1).利率
                                End If
                            End If
                            
                            ssw = 1
                        End If
                        
                        
                        Exit For
                    End If
                Next
              
             
             '*****************************************************************
             '    利息前払　利息未払　メインルーチン  2008/02/06 V182
             '*****************************************************************
               If G借入金テーブル(j).利息額 <> 0 _
                  Or Format(G借入金テーブル(j).実際年月日, "yyyy/mm/dd") _
                     = Format(w解約実行日, "yyyy/mm/dd") Then                   '10/02/02
                w対象年月 = MBA010_対象年月(CDate(G借入金テーブル(j).実際年月日))   '2008/02/06 V182
                w回目 = DateDiff("M", w基準年月, w対象年月) + 1     '2008/02/06 V182
                
                If Not IsNull(w解約実行日) Then
                    w解約締切年月日 = MBA010_締日年月日(CDate(w解約実行日))     '2008/02/06 V182
                Else
                    w解約締切年月日 = w解約実行日
                End If
                
                '***固定日数調整 08/03/05 V185
                '*初回返済年月日 or 最終返済年月日が、手打変更(wsw=1) 標準(wsw=0) 08/12/23 V189
                wsw = 0                                                             ' 08/12/23 V189
                If Format(p借入計画マスタ.初回返済年月, "yyyy/mm/dd") = _
                   Format(G借入金テーブル(j).返済予定年月, "yyyy/mm/dd") _
                   Or Format(p借入計画マスタ.最終返済年月, "yyyy/mm/dd") = _
                      Format(G借入金テーブル(j).返済予定年月, "yyyy/mm/dd") Then   ' 08/12/23 V189
                                  
                    w手打年月日 = G借入金テーブル(j).返済予定年月                   ' 08/12/23 V189
                    w手打年月日 = MXA030_翌営業年月日計算(w手打年月日, _
                                                          p借入計画マスタ.支払日, p借入計画マスタ.営業日区分) ' 08/12/23 V189
                                                          
                    If Format(w手打年月日, "yyyy/mm/dd") <> Format(G借入金テーブル(j).実際年月日, "yyyy/mm/dd") Then ' 08/12/23 V189
                        wsw = 1                                                     ' 08/12/23 V189
                    End If                                                          ' 08/12/23 V189
                End If                                                              ' 08/12/23 V189

                '10/02/04 無効にした
                'If p借入計画マスタ.実行日 <> G借入金テーブル(j).実際年月日 _
                '   And Format(w解約実行日, "yyyymmdd") <> Format(G借入金テーブル(j).実際年月日, "yyyymmdd") _
                '   And p借入計画マスタ.利息計算日数区分 = "1" _
                '   And (G借入金テーブル(j).据置X回目 = 2 Or G借入金テーブル(j).据置X回目 = 4) _
                '   And wsw = 0 Then  ' 08/12/21 V189
                '    w実際年月日 = MBA010_支払年月日算出((CDate(G借入金テーブル(j).返済予定年月)), p借入計画マスタ.支払日)
                'Else
                '    w実際年月日 = G借入金テーブル(j).実際年月日
                'End If
                
                
                If p借入計画マスタ.利息区分 = "1" Then              '2008/02/06 V182
                    If w回目 > 0 And w回目 <= 600 Then
                        p前払利息増(w回目) = p前払利息増(w回目) + G借入金テーブル(j).利息額 '2008/02/06 V182
                    End If
                    
                    '*** 10/02/04
                    If Format(G借入金テーブル(j).実際年月日, "yyyymmdd") _
                       = Format(p借入計画マスタ.実行日, "yyyymmdd") Then
                        If p借入計画マスタ.利息控除区分 = 1 Or p借入計画マスタ.利息控除区分 = 3 Then
                            w利息計算年月日 = DateAdd("d", 1, G借入金テーブル(j).利息計算年月日) '10/02/04
                        Else                                                        '10/02/04
                            w利息計算年月日 = G借入金テーブル(j).利息計算年月日     '10/02/04
                        End If                                                      '10/02/04
                    Else                                                            '10/02/04
                        w利息計算年月日 = DateAdd("d", 1, G借入金テーブル(j).利息計算年月日)  '10/02/04
                    End If                                                          '10/02/04
                    
                    If Format(G借入金テーブル(j).実際年月日, "yyyymmdd") _
                       = Format(w解約実行日, "yyyymmdd") Then               '10/02/04
                        w利息計算年月日 = G借入金テーブル(j).利息計算年月日 '10/02/04
                    End If                                                  '10/02/04
                    
                    
                    
                    'w利息計算年月日 = G借入金テーブル(j).利息計算年月日
                    Call MRB010_前払利息減計算(p借入計画マスタ, _
                                               p借入計画マスタ.実行日, _
                                               w利息計算年月日, _
                                               G借入金テーブル(j).日割日数, _
                                               G借入金テーブル(j).利息額, _
                                               w解約実行日, _
                                               w解約締切年月日, _
                                               w基準年月)     '2009/12/29 V182
                    
                                               
                Else                                                '2008/02/06 V182
                    If w回目 > 0 And w回目 <= 600 Then
                        p未払利息減(w回目) = p未払利息減(w回目) + G借入金テーブル(j).利息額 '2008/02/06 V182
                    End If
                    
                    
                    '*** 10/02/04
                    If (Format(G借入金テーブル(j).実際年月日, "yyyymmdd") _
                       = Format(p借入計画マスタ.最終返済実行日, "yyyymmdd") _
                       Or Format(G借入金テーブル(j).実際年月日, "yyyy/mm/dd") _
                          = Format(w解約実行日, "yyyy/mm/dd")) _
                       And _
                         (p借入計画マスタ.利息控除区分 = 2 Or p借入計画マスタ.利息控除区分 = 3) Then
                        w利息計算年月日 = DateAdd("d", -1, G借入金テーブル(j).利息計算年月日)
                    Else                    '10/02/04
                        w利息計算年月日 = G借入金テーブル(j).利息計算年月日     '10/02/04
                    End If
                    
                    Call MRB010_未払利息増計算(p借入計画マスタ, _
                                               p借入計画マスタ.実行日, _
                                               w利息計算年月日, _
                                               G借入金テーブル(j).日割日数, _
                                               G借入金テーブル(j).利息額, _
                                               w解約実行日, _
                                               w解約締切年月日, _
                                               w基準年月)     '2008/02/06 V182
                    
                End If                                              '2008/02/06 V182
                
               End If
               
             Next
             
             '***終了年月算出
             If Not IsNull(w解約実行日) Then                    '2008/02/07 V182
                w終了年月 = w解約実行日                         '2008/02/07 V182
             Else                                               '2008/02/07 V182
                w終了年月 = p借入計画マスタ.最終返済実行日      '2008/02/07 V182
             End If                                             '2008/02/07 V182
             
             w終了年月 = MBA010_対象年月(CDate(w終了年月))      '2008/02/07 V182
             
             w開始回目 = DateDiff("M", w基準年月, w基準年月) + 1 '2008/02/07 V182
             w終了回目 = DateDiff("M", w基準年月, w終了年月) + 1 '2008/02/07 V182
             
             '**************************************************************
             '     前払利息　未払利息　集計セット
             '**************************************************************
             For j = w開始回目 To w終了回目                     '2008/02/07 V182
                w対象年月 = DateAdd("M", j - 1, w基準年月)      '2008/02/07 V182
                If j = 1 Then                                   '2008/02/07 V182
                    p前払利息残(j) = p前払利息増(j) - p前払利息減(j) '2008/02/07V182
                    p未払利息残(j) = p未払利息増(j) - p未払利息減(j) '2008/02/07V182
                Else                                            '2008/02/07 V182
                    p前払利息残(j) = p前払利息残(j - 1) + p前払利息増(j) - p前払利息減(j) '2008/02/07 V182
                    p未払利息残(j) = p未払利息残(j - 1) + p未払利息増(j) - p未払利息減(j) '2008/02/07V182
                End If                                          '2008/02/07 V182
                
                For k = 1 To wcnt                               '2008/02/07 V182
                    If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                        And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then '2008/02/07 V182
                            
                        w前払利息増(k) = w前払利息増(k) + p前払利息増(j) '2008/02/07 V182
                        w前払利息減(k) = w前払利息減(k) + p前払利息減(j) '2008/02/07 V182
                        w未払利息増(k) = w未払利息増(k) + p未払利息増(j) '2008/02/07 V182
                        w未払利息減(k) = w未払利息減(k) + p未払利息減(j) '2008/02/07 V182
                        
                        If Format(w対象年月, "yyyymmdd") = Format(w年月(k), "yyyymmdd") Then '2008/02/07 V182
                            w前払利息(k) = p前払利息残(j)       '2008/02/07 V182
                            w未払利息(k) = p未払利息残(j)       '2008/02/07 V182
                        End If                                  '2008/02/07 V182
                    End If                                      '2008/02/07 V182
                Next                                            '2008/02/07 V182
             Next                                               '2008/02/07 V182
             
             For k = 1 To wcnt
                If k > 1 Then
                    w前月残高(k) = w残高(k - 1)
                End If
                
                w残高(k) = w前月残高(k) + w融資(k) - w元金(k) - w解約(k)    ' 07/02/18 V180
                
             Next
             
          

               
             For k = 1 To wcnt
                If w解約(k) <> 0 Then               '11/06/17 V200
                    w元金(k) = w元金(k) + w解約(k)  '11/06/17 V200
                    w返済(k) = w返済(k) + w解約(k)  '11/06/17 V200
                End If                              '11/06/17 V200
                
                w融資合計 = w融資合計 + w融資(k)
                w元金合計 = w元金合計 + w元金(k)
                w利息合計 = w利息合計 + w利息(k)
                w返済合計 = w返済合計 + w返済(k)
                w解約合計 = w解約合計 + w解約(k)
                w保証合計 = w保証合計 + w保証(k)
                w手数料合計 = w手数料合計 + w手数料(k)                  '08/12/09 V189
                w前払利息増合計 = w前払利息増合計 + w前払利息増(k)      '2008/02/07 V182
                w前払利息減合計 = w前払利息減合計 + w前払利息減(k)      '2008/02/07 V182
                w未払利息増合計 = w未払利息増合計 + w未払利息増(k)      '2008/02/07 V182
                w未払利息減合計 = w未払利息減合計 + w未払利息減(k)      '2008/02/07 V182
                
             Next
             
             
             '***利率= 0 の時　調整処理
             w現在利率 = w直前利率
             For k = 1 To wcnt
                If w融資(k) <> 0 Or w元金(k) <> 0 Or w利息(k) <> 0 _
                                 Or w解約(k) <> 0 Or w残高(k) <> 0 Then
                    If w利率(k) = 0 Then
                        w利率(k) = w現在利率
                    End If
                    
                    w現在利率 = w利率(k)
                End If
             Next
             
             
             
             '***決算用の利率調整 2012/02/24
             For k = 1 To wcnt
                If (w残高(k) <> 0 Or w元金(k) <> 0 Or w利息(k) <> 0) And w利率(k) = 0 Then
                    If k = 1 Then
                        w利率(k) = w利率(k + 1)
                    Else
                        w利率(k) = w利率(k - 1)
                    End If
                End If
             Next
             
             
             w残高合計 = w残高(wcnt)
             w前払利息合計 = w前払利息(wcnt)                            '2008/02/07 V182
             w未払利息合計 = w未払利息(wcnt)                            '2008/02/07 V182
                
             If w融資合計 = 0 And w元金合計 = 0 And w利息合計 = 0 And _
                w返済合計 = 0 And w解約合計 = 0 And w保証合計 = 0 And _
                w前払利息増合計 = 0 And w前払利息減合計 = 0 And _
                w未払利息増合計 = 0 And w未払利息減合計 = 0 And _
                w残高合計 = 0 Then
             Else
                 If FLG_Mdata = True Then
                    wRs2.AddNew
                        wRs2("借入番号") = w借入番号
                        wRs2("融資合計") = w融資合計
                        wRs2("元金合計") = w元金合計
                        wRs2("利息合計") = w利息合計
                        wRs2("返済合計") = w返済合計
                        wRs2("解約合計") = w解約合計
                        wRs2("保証合計") = w保証合計
                        wRs2("手数料合計") = w手数料合計            ' 08/12/09 V189
                        
                        wRs2("初期手数料合計") = w初期手数料合計
                        wRs2("元金手数料合計") = w元金手数料合計
                        wRs2("利息手数料合計") = w利息手数料合計
                        
                        wRs2("残高合計") = w残高合計
    
                        wRs2("前払利息増合計") = w前払利息増合計    '2008/02/07 V182
                        wRs2("前払利息減合計") = w前払利息減合計    '2008/02/07 V182
                        wRs2("前払利息合計") = w前払利息合計        '2008/02/07 V182
                        wRs2("未払利息増合計") = w未払利息増合計    '2008/02/07 V182
                        wRs2("未払利息減合計") = w未払利息減合計    '2008/02/07 V182
                        wRs2("未払利息合計") = w未払利息合計        '2008/02/07 V182
    
    
                        For k = 1 To wcnt
                            wRs2("融資_" + CStr(Format(k, "00"))) = w融資(k)
                            wRs2("元金_" + CStr(Format(k, "00"))) = w元金(k)
                            wRs2("利息_" + CStr(Format(k, "00"))) = w利息(k)
                            wRs2("返済_" + CStr(Format(k, "00"))) = w返済(k)
                            wRs2("解約_" + CStr(Format(k, "00"))) = w解約(k)
                            wRs2("保証_" + CStr(Format(k, "00"))) = w保証(k)
                            wRs2("手数料_" + CStr(Format(k, "00"))) = w手数料(k)        ' 08/12/09 V189
                            
                            wRs2("初期手数料_" + CStr(Format(k, "00"))) = w初期手数料(k)
                            wRs2("元金手数料_" + CStr(Format(k, "00"))) = w元金手数料(k)
                            wRs2("利息手数料_" + CStr(Format(k, "00"))) = w利息手数料(k)
                            
                            wRs2("残高_" + CStr(Format(k, "00"))) = w残高(k)
    
                            wRs2("前払利息増_" + CStr(Format(k, "00"))) = w前払利息増(k) '2008/02/07 V182
                            wRs2("前払利息減_" + CStr(Format(k, "00"))) = w前払利息減(k) '2008/02/07 V182
                            wRs2("前払利息_" + CStr(Format(k, "00"))) = w前払利息(k)     '2008/02/07 V182
                            wRs2("未払利息増_" + CStr(Format(k, "00"))) = w未払利息増(k) '2008/02/07 V182
                            wRs2("未払利息減_" + CStr(Format(k, "00"))) = w未払利息減(k) '2008/02/07 V182
                            wRs2("未払利息_" + CStr(Format(k, "00"))) = w未払利息(k)     '2008/02/07 V182
                            
                            wRs2("利率_" + CStr(Format(k, "00"))) = w利率(k)             '11/02/17
                            
                        Next
    
                    wRs2.Update
                
                 Else
                        
                    '2010/06/18
                    p推移(wiCnt).借入番号 = w借入番号
                    p推移(wiCnt).融資合計 = w融資合計
                    p推移(wiCnt).元金合計 = w元金合計
                    p推移(wiCnt).利息合計 = w利息合計
                    p推移(wiCnt).返済合計 = w返済合計
                    p推移(wiCnt).解約合計 = w解約合計
                    p推移(wiCnt).保証合計 = w保証合計
                    p推移(wiCnt).手数料合計 = w手数料合計
                    
                    p推移(wiCnt).初期手数料合計 = w初期手数料合計
                    p推移(wiCnt).元金手数料合計 = w元金手数料合計
                    p推移(wiCnt).利息手数料合計 = w利息手数料合計
                    
                    p推移(wiCnt).残高合計 = w残高合計

                    p推移(wiCnt).前払利息増合計 = w前払利息増合計
                    p推移(wiCnt).前払利息減合計 = w前払利息減合計
                    p推移(wiCnt).前払利息合計 = w前払利息合計
                    p推移(wiCnt).未払利息増合計 = w未払利息増合計
                    p推移(wiCnt).未払利息減合計 = w未払利息減合計
                    p推移(wiCnt).未払利息合計 = w未払利息合計

                    For k = 1 To wcnt
                        p推移(wiCnt).融資(k) = w融資(k)
                        p推移(wiCnt).元金(k) = w元金(k)
                        p推移(wiCnt).利息(k) = w利息(k)
                        p推移(wiCnt).返済(k) = w返済(k)
                        p推移(wiCnt).解約(k) = w解約(k)
                        p推移(wiCnt).保証(k) = w保証(k)
                        p推移(wiCnt).手数料(k) = w手数料(k)        ' 08/12/09 V189
                        
                        p推移(wiCnt).初期手数料(k) = w初期手数料(k)
                        p推移(wiCnt).元金手数料(k) = w元金手数料(k)
                        p推移(wiCnt).利息手数料(k) = w利息手数料(k)
                        
                        p推移(wiCnt).残高(k) = w残高(k)

                        p推移(wiCnt).前払利息増(k) = w前払利息増(k) '2008/02/07 V182
                        p推移(wiCnt).前払利息減(k) = w前払利息減(k) '2008/02/07 V182
                        p推移(wiCnt).前払利息(k) = w前払利息(k)     '2008/02/07 V182
                        p推移(wiCnt).未払利息増(k) = w未払利息増(k) '2008/02/07 V182
                        p推移(wiCnt).未払利息減(k) = w未払利息減(k) '2008/02/07 V182
                        p推移(wiCnt).未払利息(k) = w未払利息(k)     '2008/02/07 V182
                        
                        p推移(wiCnt).利率(k) = w利率(k)             '11/02/17
                        
                    Next

                    wiCnt = wiCnt + 1
                 End If
             End If
                
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing

    If FLG_Mdata = True Then
        wRs2.Close
        Set wRs = Nothing
    End If
'
    If FLG_Mdata <> True Then
        wstr2 = ""
        wstr2 = wstr2 + "Select * From DCDA010_借入残高推移表結果"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
        
            For wiCnt = 0 To 9999
            
                If p推移(wiCnt).借入番号 = "" Then
                    Exit For
                End If
                
                wRs2.AddNew
                    wRs2("借入番号") = p推移(wiCnt).借入番号
                    wRs2("融資合計") = p推移(wiCnt).融資合計
                    wRs2("元金合計") = p推移(wiCnt).元金合計
                    wRs2("利息合計") = p推移(wiCnt).利息合計
                    wRs2("返済合計") = p推移(wiCnt).返済合計
                    wRs2("解約合計") = p推移(wiCnt).解約合計
                    wRs2("保証合計") = p推移(wiCnt).保証合計
                    wRs2("手数料合計") = p推移(wiCnt).手数料合計
                    
                    wRs2("初期手数料合計") = p推移(wiCnt).初期手数料合計
                    wRs2("元金手数料合計") = p推移(wiCnt).元金手数料合計
                    wRs2("利息手数料合計") = p推移(wiCnt).利息手数料合計
                    
                    wRs2("残高合計") = p推移(wiCnt).残高合計
        
                    wRs2("前払利息増合計") = p推移(wiCnt).前払利息増合計
                    wRs2("前払利息減合計") = p推移(wiCnt).前払利息減合計
                    wRs2("前払利息合計") = p推移(wiCnt).前払利息合計
                    wRs2("未払利息増合計") = p推移(wiCnt).未払利息増合計
                    wRs2("未払利息減合計") = p推移(wiCnt).未払利息減合計
                    wRs2("未払利息合計") = p推移(wiCnt).未払利息合計
        
                    For k = 1 To wcnt
                        wRs2("融資_" + CStr(Format(k, "00"))) = p推移(wiCnt).融資(k)
                        wRs2("元金_" + CStr(Format(k, "00"))) = p推移(wiCnt).元金(k)
                        wRs2("利息_" + CStr(Format(k, "00"))) = p推移(wiCnt).利息(k)
                        wRs2("返済_" + CStr(Format(k, "00"))) = p推移(wiCnt).返済(k)
                        wRs2("解約_" + CStr(Format(k, "00"))) = p推移(wiCnt).解約(k)
                        wRs2("保証_" + CStr(Format(k, "00"))) = p推移(wiCnt).保証(k)
                        wRs2("手数料_" + CStr(Format(k, "00"))) = p推移(wiCnt).手数料(k)
                        
                        wRs2("初期手数料_" + CStr(Format(k, "00"))) = p推移(wiCnt).初期手数料(k)
                        wRs2("元金手数料_" + CStr(Format(k, "00"))) = p推移(wiCnt).元金手数料(k)
                        wRs2("利息手数料_" + CStr(Format(k, "00"))) = p推移(wiCnt).利息手数料(k)
                        
                        wRs2("残高_" + CStr(Format(k, "00"))) = p推移(wiCnt).残高(k)
        
                        wRs2("前払利息増_" + CStr(Format(k, "00"))) = p推移(wiCnt).前払利息増(k)
                        wRs2("前払利息減_" + CStr(Format(k, "00"))) = p推移(wiCnt).前払利息減(k)
                        wRs2("前払利息_" + CStr(Format(k, "00"))) = p推移(wiCnt).前払利息(k)
                        wRs2("未払利息増_" + CStr(Format(k, "00"))) = p推移(wiCnt).未払利息増(k)
                        wRs2("未払利息減_" + CStr(Format(k, "00"))) = p推移(wiCnt).未払利息減(k)
                        wRs2("未払利息_" + CStr(Format(k, "00"))) = p推移(wiCnt).未払利息(k)
                        
                        wRs2("利率_" + CStr(Format(k, "00"))) = p推移(wiCnt).利率(k)    '11/02/17
                        
                    Next
        
                wRs2.Update
        
            Next wiCnt
            
        wRs2.Close
        Set wRs2 = Nothing
    End If
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MRB010_標準入力借入残高表固定日数_ERR:
    pERR_MES = pPROGRAM_ID + "/ MRB010_標準入力借入残高表固定日数() でエラー" + vbCrLf + vbCrLf + _
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
' MRB010_標準入力未払前払
'------------------------------------------------
Public Sub MRB010_標準入力未払前払(pTbl As String, p借入番号 As String)
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String, wWhere As String
    
    Dim wiCnt As Integer
    Dim p推移(9999) As MRB010_借入金推移表
    Dim FLG_Mdata As Boolean
    
    Dim p借入計画マスタ As MAA910_借入金                                       '5/10/8 V129
    Dim w銀行マスタ As MAA030_銀行
    Dim w借入金 As MAA910_借入金
    
    Dim j As Integer, k As Integer
    Dim w千円単位 As Integer
    Dim w年 As Integer, w月 As Integer, w日 As Integer                         '5/8/18 V129
    Dim ww月 As Integer                                                        '5/9/2 V129
    Dim w間隔 As Integer, w回数 As Integer, wcnt As Integer
    
    '5/9/12 V129 解約の時 1回前の年月との差＞＝２　1回前の融資残更新
    Dim w月差 As Integer
    Dim w処理 As Integer                                                       '5/9/9 V129
    
    Dim w残高算出 As Double                                                  '5/8/18 V129
    Dim w融資残高 As Double                                                  '5/9/9 V129
    Dim w前月残合計 As Double, w前月残高(12) As Double                          ' 07/02/09 V180
    Dim w融資合計 As Double, w融資(12) As Double
    Dim w元金合計 As Double, w元金(12) As Double
    Dim w利息合計 As Double, w利息(12) As Double
    Dim w返済合計 As Double, w返済(12) As Double
    Dim w解約合計 As Double, w解約(12) As Double
    Dim w残高合計 As Double, w残高(12) As Double
    Dim w保証合計 As Double, w保証(12) As Double
    Dim w手数料合計 As Double, w手数料(12) As Double        ' 08/12/09 V189
    
    Dim w初期手数料合計 As Double, w初期手数料(12) As Double
    Dim w元金手数料合計 As Double, w元金手数料(12) As Double
    Dim w利息手数料合計 As Double, w利息手数料(12) As Double
    
    
    Dim w前払利息増合計 As Double, w前払利息増(12)          '2008/02/06 V182
    Dim w前払利息減合計 As Double, w前払利息減(12)          '2008/02/06 V182
    Dim w前払利息合計 As Double, w前払利息(12)              '2008/02/06 V182
    Dim w未払利息増合計 As Double, w未払利息増(12)          '2008/02/06 V182
    Dim w未払利息減合計 As Double, w未払利息減(12)          '2008/02/06 V182
    Dim w未払利息合計 As Double, w未払利息(12)              '2008/02/06 V182
    Dim w基準年月 As Date                                   '2008/02/06 V182
    Dim w終了年月 As Variant                                '2008/02/06 V182
    Dim w開始回目 As Integer                                '2008/02/07 V182
    Dim w終了回目 As Integer                                '2008/02/07 V182
    Dim w回目 As Integer                                    '2008/02/06 V182
    Dim w解約締切年月日 As Variant                          '2008/02/06 V182
    Dim w利息計算年月日 As Variant                          '10/02/04
        
    Dim wd01 As Date
    Dim w実際年月 As Date
    Dim w実際年月日 As Variant                              '08/03/05 V185
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
    Dim w借入番号 As String, w借入計画番号 As String, w金融リストラ As String
    
    Dim w解約判定 As String
    Dim ws基本 As String                                                       '5/8/30 V129
    Dim wベンチャ As String, wベンチャcode As String
    
    Dim w借入貸付 As String                                                     ' 07/02/09 V180
    Dim w前月残 As Double                                                       ' 07/02/09 V180
    
    Dim w手打年月日 As Date                                                     ' 08/12/23 V189
    Dim wsw As Integer                                                          ' 08/12/23 V189
    
    '5/10/9 V129 支店への貸付　リストラ番号
    Dim w期首借入年度 As String, w支店貸付 As String, w全社借入 As String
'    Dim wsTbl As String
'
    On Error GoTo MRB010_標準入力未払前払_ERR
'
    ' -----------------------------------------
    '       借入金マスタより DCDA010_借入残高推移表結果　作成
    ' -----------------------------------------
    'w開始年 = GRpt.テキスト_01
    w開始年 = 22                    '11/02/14
    w借入計画番号 = GRpt.借入
    w金融リストラ = GRpt.金融
    w千円単位 = GRpt.千円単位

    If w千円単位 = 1 Then
        w分母 = "1000"
    Else
        w分母 = "1"
    End If
    
    ReDim G利息未払前払テーブル(0)
    
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
    GRpt.推移 = "月次"                          '11/02/14
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
    
    
    '** ワークファイル 削除 **
    wstr2 = ""
    wstr2 = wstr2 + "Delete * From DCDA010_借入残高推移表結果"
    GDb.Execute wstr2
    
    
    FLG_Mdata = False '通常はデータ一括書込
    wiCnt = 0
'
    wstr = ""
    wstr = wstr + "Select * From " & pTbl
    wstr = wstr + " Where 借入番号 = '" & p借入番号 & "'"
    wstr = wstr + " And 手入力区分 = 0"                   ' 07/02/09 V180
    
    '**ベンチャーの場合
    If (wベンチャcode < "a" Or wベンチャcode > "z") And w借入計画番号 <> "" Then
        wstr = wstr + " And ("                               ' 07/02/09 V180
        
        If w金融リストラ <> "" Then
            wstr = wstr + " (金融リストラ番号='" & w金融リストラ & "' And 取消フラグ=0) Or  "
        End If
        
        wstr = wstr + " (借入計画番号='" & w借入計画番号 & "' And 借入計画番号 <> ''"
        
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
        wstr = wstr + " Or (借入計画番号='" & w借入計画番号 & "' And 借入計画番号 <> ''"
        
        If G基本情報.企業区分 = "連結本部" Then         '06/09/22 V170
            wstr = wstr + " And Format(実行日,'yyyymmdd') > '" & Format(Gコントロール.最終実績年月, "yyyymmdd") & "'"
        End If                                    '17/8/17 V128
        
        wstr = wstr + " And 取消フラグ=0)"
        
        If G基本情報.企業区分 = "連結本部" Or G基本情報.企業区分 = "連結子会社" Or G基本情報.企業区分 = "全社" Then
            wstr = wstr + " Or (借入計画番号='" & w借入計画番号 & "' And sm区分=1 And 取消フラグ=0)"
        End If
    
    End If
        wstr = wstr + ")"

    'If pTbl2 <> "" Then
    '    wstr = wstr + " UNION Select * From " & pTbl2
    '    wstr = wstr + " Where 手入力区分 = 1"
    '    wstr = wstr + " And ((借入計画番号='' And sm区分 = 0 And 取消フラグ = 0)"
    '    wstr = wstr + " Or (借入計画番号 = '" & w借入計画番号 & "' And 借入計画番号 <> ''"
    '    wstr = wstr + " And sm区分 = 0 And 取消フラグ = 0))"
    'End If
        
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.RecordCount >= 10000 Then
        FLG_Mdata = True
        
        wstr2 = ""
        wstr2 = wstr2 + "Select * From DCDA010_借入残高推移表結果"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
    End If
        
        Do Until wRs.eof
            p借入計画マスタ = MBD010_借入データセット(wRs)      '5/10/8 V129
         
            '** 借入金テーブル セット **
            
            Call MBD010_借入金テーブル作成(w金融リストラ, p借入計画マスタ)    ' 07/02/18 V180
            
            w借入番号 = p借入計画マスタ.借入番号                '5/10/8 V129
            For j = 1 To 12
                w前月残高(j) = 0                                ' 07/02/09 V180
                w融資(j) = 0
                w元金(j) = 0
                w利息(j) = 0
                w返済(j) = 0
                w解約(j) = 0
                w残高(j) = 0
                w保証(j) = 0
                w手数料(j) = 0                      ' 08/12/09 V189
                
                w初期手数料(j) = 0
                w元金手数料(j) = 0
                w利息手数料(j) = 0
                
                w前払利息増(j) = 0                  '2008/02/06 V182
                w前払利息減(j) = 0                  '2008/02/06 V182
                w前払利息(j) = 0                    '2008/02/06 V182
                w未払利息増(j) = 0                  '2008/02/06 V182
                w未払利息減(j) = 0                  '2008/02/06 V182
                w未払利息(j) = 0                    '2008/02/06 V182
                
            Next
            
            w融資合計 = 0
            w元金合計 = 0
            w利息合計 = 0
            w返済合計 = 0
            w解約合計 = 0
            w残高合計 = 0
            w保証合計 = 0
            w手数料合計 = 0                         ' 08/12/09 V189
            
            w初期手数料合計 = 0
            w元金手数料合計 = 0
            w利息手数料合計 = 0
            
            w前払利息増合計 = 0                     '2008/02/06 V182
            w前払利息減合計 = 0                     '2008/02/06 V182
            w前払利息合計 = 0                       '2008/02/06 V182
            w未払利息増合計 = 0                     '2008/02/06 V182
            w未払利息減合計 = 0                     '2008/02/06 V182
            w未払利息合計 = 0                       '2008/02/06 V182
            
            For w回目 = 1 To 600                    '2008/02/06 V182
                p前払利息増(w回目) = 0              '2008/02/06 V182
                p前払利息減(w回目) = 0              '2008/02/06 V182
                p前払利息残(w回目) = 0              '2008/02/06 V182
                p未払利息増(w回目) = 0              '2008/02/06 V182
                p未払利息減(w回目) = 0              '2008/02/06 V182
                p未払利息残(w回目) = 0              '2008/02/06 V182
            Next                                    '2008/02/06 V182
            
            
            '
            
            ' =========================================
            '                 初期設定
            ' =========================================
            '期首前残設定
            
            w手入力区分 = p借入計画マスタ.手入力区分            '11/02/16
            
            '***
             'Call MBD010_借入金入力明細Read(p借入計画マスタ.借入番号, p借入計画マスタ.借入貸付)  ' 07/02/09 V180
            
            w対象年月 = DateAdd("m", 1, w年月(0))                   ' 07/02/09 V180
            'w前月残 = MBD010_借入金手入力残高(p借入計画マスタ, 0, w対象年月) ' 07/02/09 V180
            w前月残 = MBD010_借入金標準入力残高(p借入計画マスタ, w金融リストラ, 0, w対象年月, G基本情報.借入金管理区分) '07/02/26 V180
            w前月残高(1) = w前月残                                  ' 07/02/09 V180
            w対象年月 = MBA010_対象年月(CDate(p借入計画マスタ.実行日))  ' 07/02/09 V180
            
            w基準年月 = MBA010_対象年月(CDate(p借入計画マスタ.実行日))      '2008/02/06 V182
            
              
             For k = 1 To wcnt                                                               '5/10/8 V129
                 If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                     And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then '5/10/8 V129
                     w融資(k) = w融資(k) + p借入計画マスタ.融資金額                             '5/10/8 V129
                     Exit For                                                                '5/10/9 V129
                 End If                                                                      '5/10/8 V129
             Next                                                                            '5/10/8 V129
                 
                 
             w前払利息残 = 0                                        '11/02/11
             w未払利息残 = 0                                        '11/02/11
             
                 
                 
             For j = 1 To UBound(G借入金テーブル)                   ' 07/02/18 V180
             
                Call MBA010_借入金年月算出(G借入金テーブル(j).返済予定年月, _
                    G借入金テーブル(j).実際年月日, p借入計画マスタ.支払日)  ' 07/02/12 V180
                    
                If G基本情報.借入金管理区分 = XMXA020_区分("借入金管理区分", "管理用") Then '07/02/18 V180
                    w対象年月 = MBA010_対象年月((CDate(G管理年月)))             ' 07/02/18 V180
                Else                                                            ' 07/02/180V180
                    w対象年月 = MBA010_対象年月((CDate(G実績年月)))             ' 07/02/18 V180
                End If                                                          ' 07/02/180V180
                
                '解約算出
                If w金融リストラ > "" _
                        And w金融リストラ = p借入計画マスタ.金融リストラ番号 Then  ' 07/02/18 V180
                        w解約実行日 = p借入計画マスタ.金融解約実行日
                Else
                        w解約実行日 = p借入計画マスタ.解約実行日
                End If
                
                If Format(w解約実行日, "yyyymmdd") = _
                            Format(G借入金テーブル(j).実際年月日, "yyyymmdd") Then  ' 07/02/18 V180
                    w対象年月 = MBA010_対象年月((CDate(G実績年月)))             ' 07/02/18 V180
                End If                                                          ' 07/02/18 V180
                        
             
                'w対象年月 = MBA010_対象年月(CDate(G借入金テーブル (J).借入返済年月日))
                For k = 1 To wcnt
                    If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                       And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then
                        w元金(k) = w元金(k) + G借入金テーブル(j).元金額     ' 07/02/18 V180
                        w利息(k) = w利息(k) + G借入金テーブル(j).利息額     ' 07/02/18 V180
                        w返済(k) = w返済(k) + G借入金テーブル(j).返済金額   ' 07/02/18 V180
                        w保証(k) = w保証(k) + G借入金テーブル(j).保証料     ' 07/02/18 V180
                        w手数料(k) = w手数料(k) + G借入金テーブル(j).手数料 ' 08/12009 V189
                        
                        If Format(w解約実行日, "yyyymmdd") = _
                            Format(G借入金テーブル(j).実際年月日, "yyyymmdd") Then  ' 07/02/18 V180
                            w解約(k) = w解約(k) + G借入金テーブル(j).融資残高       ' 07/02/18 V180
                        End If                                                  ' 07/02/18 V180
                        
                        Exit For
                    End If
                Next
              
             
             '*****************************************************************
             '    利息前払　利息未払　メインルーチン  2008/02/06 V182
             '*****************************************************************
               If G借入金テーブル(j).利息額 <> 0 _
                  Or Format(G借入金テーブル(j).実際年月日, "yyyy/mm/dd") _
                     = Format(w解約実行日, "yyyy/mm/dd") Then                   '10/02/02
                w対象年月 = MBA010_対象年月(CDate(G借入金テーブル(j).実際年月日))   '2008/02/06 V182
                w回目 = DateDiff("M", w基準年月, w対象年月) + 1     '2008/02/06 V182
                
                If Not IsNull(w解約実行日) Then
                    w解約締切年月日 = MBA010_締日年月日(CDate(w解約実行日))     '2008/02/06 V182
                Else
                    w解約締切年月日 = w解約実行日
                End If
                
                '***固定日数調整 08/03/05 V185
                '*初回返済年月日 or 最終返済年月日が、手打変更(wsw=1) 標準(wsw=0) 08/12/23 V189
                wsw = 0                                                             ' 08/12/23 V189
                If Format(p借入計画マスタ.初回返済年月, "yyyy/mm/dd") = _
                   Format(G借入金テーブル(j).返済予定年月, "yyyy/mm/dd") _
                   Or Format(p借入計画マスタ.最終返済年月, "yyyy/mm/dd") = _
                      Format(G借入金テーブル(j).返済予定年月, "yyyy/mm/dd") Then   ' 08/12/23 V189
                                  
                    w手打年月日 = G借入金テーブル(j).返済予定年月                   ' 08/12/23 V189
                    w手打年月日 = MXA030_翌営業年月日計算(w手打年月日, _
                                                          p借入計画マスタ.支払日, p借入計画マスタ.営業日区分) ' 08/12/23 V189
                                                          
                    If Format(w手打年月日, "yyyy/mm/dd") <> Format(G借入金テーブル(j).実際年月日, "yyyy/mm/dd") Then ' 08/12/23 V189
                        wsw = 1                                                     ' 08/12/23 V189
                    End If                                                          ' 08/12/23 V189
                End If                                                              ' 08/12/23 V189

                '10/02/04 無効にした
                'If p借入計画マスタ.実行日 <> G借入金テーブル(j).実際年月日 _
                '   And Format(w解約実行日, "yyyymmdd") <> Format(G借入金テーブル(j).実際年月日, "yyyymmdd") _
                '   And p借入計画マスタ.利息計算日数区分 = "1" _
                '   And (G借入金テーブル(j).据置X回目 = 2 Or G借入金テーブル(j).据置X回目 = 4) _
                '   And wsw = 0 Then  ' 08/12/21 V189
                '    w実際年月日 = MBA010_支払年月日算出((CDate(G借入金テーブル(j).返済予定年月)), p借入計画マスタ.支払日)
                'Else
                '    w実際年月日 = G借入金テーブル(j).実際年月日
                'End If
                
                
                
                '***head部　標準SET
                w利息未払前払.銀行番号 = p借入計画マスタ.銀行番号   '11/02/11
                w利息未払前払.借入番号 = p借入計画マスタ.借入番号   '11/02/11
                w利息未払前払.利息区分 = p借入計画マスタ.利息区分   '11/02/11
                w利息未払前払.利息計算日数区分 = p借入計画マスタ.利息計算日数区分   '11/02/11
                
                
                
                If p借入計画マスタ.利息区分 = "1" Then              '2008/02/06 V182
                    If w回目 > 0 And w回目 <= 600 Then
                        p前払利息増(w回目) = p前払利息増(w回目) + G借入金テーブル(j).利息額 '2008/02/06 V182
                    End If
                    
                    '*** 10/02/04
                    If Format(G借入金テーブル(j).実際年月日, "yyyymmdd") _
                       = Format(p借入計画マスタ.実行日, "yyyymmdd") Then
                        If p借入計画マスタ.利息控除区分 = 1 Or p借入計画マスタ.利息控除区分 = 3 Then
                            w利息計算年月日 = DateAdd("d", 1, G借入金テーブル(j).利息計算年月日) '10/02/04
                        Else                                                        '10/02/04
                            w利息計算年月日 = G借入金テーブル(j).利息計算年月日     '10/02/04
                        End If                                                      '10/02/04
                    Else                                                            '10/02/04
                        w利息計算年月日 = DateAdd("d", 1, G借入金テーブル(j).利息計算年月日)  '10/02/04
                    End If                                                          '10/02/04
                    
                    If Format(G借入金テーブル(j).実際年月日, "yyyymmdd") _
                       = Format(w解約実行日, "yyyymmdd") Then               '10/02/04
                        w利息計算年月日 = G借入金テーブル(j).利息計算年月日 '10/02/04
                    End If                                                  '10/02/04
                    
                    
                    '***ITEM部　標準前払い増
                    w利息未払前払.返済年月日 = G借入金テーブル(j).実際年月日    '11/02/11
                    w利息未払前払.月毎NO = 1                                    '11/02/11
                    w利息未払前払.元金額 = G借入金テーブル(j).元金額                  '11/02/11
                    w利息未払前払.融資残高 = G借入金テーブル(j).融資残高            '11/02/11
                    
                    '*うち入れ 利息計算対象額
                    If G借入金テーブル(j).据置x回目 = 3 Then            '2012/03/01
                        w利息未払前払.利息計算対象額 = G借入金テーブル(j).元金額    '2012/03/01
                    Else                                                '2012/03/01
                        w利息未払前払.利息計算対象額 = w利息未払前払.融資残高       '11/02/16
                    End If                                              '2012/03/01
                    
                    w利息未払前払.利息額増 = G借入金テーブル(j).利息額              '11/02/11
                    w利息未払前払.利息額減 = 0                                  '11/02/11
                    w前払利息残 = w前払利息残 + G借入金テーブル(j).利息額           '11/02/11
                    w利息未払前払.利息残高 = w前払利息残                        '11/02/11
                    w利息未払前払.日割日数 = G借入金テーブル(j).日割日数        '11/02/11
                    w利息未払前払.利率 = G借入金テーブル(j).利率                    '11/02/11
                    
                    If G借入金テーブル(j).実際年月日 = w解約実行日 And p借入計画マスタ.利息控除区分 <= 1 Then         '2012/03/21
                        w利息未払前払.開始年月日 = DateAdd("d", 1, w利息計算年月日)
                    Else                                                        '11/02/15
                        w利息未払前払.開始年月日 = w利息計算年月日                  '11/02/11
                    End If                                                      '11/02/15
                    'w利息未払前払.開始年月日 = w利息計算年月日                  '2012/03/21
                    
                    If G借入金テーブル(j).日割日数 < 0 Then                         '11/02/11
                        w利息未払前払.終了年月日 = DateAdd("d", -G借入金テーブル(j).日割日数 - 1, w利息未払前払.開始年月日) '11/02/11
                    Else                                                        '11/02/11
                        w利息未払前払.終了年月日 = DateAdd("d", G借入金テーブル(j).日割日数 - 1, w利息未払前払.開始年月日) '11/02/11
                    End If                                                      '11/02/11
                    
                    w利息未払前払.据置x回目 = G借入金テーブル(j).据置x回目  '2012/03/01
                    
                    
                    w利息未払前払.利息期間対象日数 = G借入金テーブル(j).日割日数    '2014/08/26
                    w利息未払前払.利息期間対象額 = G借入金テーブル(j).利息額        '2014/08/26
                    w利息未払前払.利息調整F = 0                                     '2014/08/26
                    
                    Call MBD010_利息未払前払Write(w利息未払前払, p借入計画マスタ)   '2016/09/26
                    
                   
                    
                    
                    
                    'w利息計算年月日 = G借入金テーブル(j).利息計算年月日
                    Call MRB010_前払利息減計算明細(p借入計画マスタ, _
                                               p借入計画マスタ.実行日, _
                                               p借入計画マスタ.利息控除区分, _
                                               w利息計算年月日, _
                                               G借入金テーブル(j).日割日数, _
                                               G借入金テーブル(j).利息額, _
                                               w解約実行日, _
                                               w解約締切年月日, _
                                               w基準年月)     '2009/12/29 V182
                    
                                               
                Else                                                '2008/02/06 V182
                    If w回目 > 0 And w回目 <= 600 Then
                        p未払利息減(w回目) = p未払利息減(w回目) + G借入金テーブル(j).利息額 '2008/02/06 V182
                    End If
                    
                    
                    '*** 10/02/04
                    If (Format(G借入金テーブル(j).実際年月日, "yyyymmdd") _
                       = Format(p借入計画マスタ.最終返済実行日, "yyyymmdd") _
                       Or Format(G借入金テーブル(j).実際年月日, "yyyy/mm/dd") _
                          = Format(w解約実行日, "yyyy/mm/dd")) _
                       And _
                         (p借入計画マスタ.利息控除区分 = 2 Or p借入計画マスタ.利息控除区分 = 3) Then
                        w利息計算年月日 = DateAdd("d", -1, G借入金テーブル(j).利息計算年月日)
                    Else                    '10/02/04
                        w利息計算年月日 = G借入金テーブル(j).利息計算年月日     '10/02/04
                    End If
                    
                                                 
                    
                    '***ITEM部　標準未払い減
                    w利息未払前払.返済年月日 = G借入金テーブル(j).実際年月日    '11/02/11
                    w利息未払前払.月毎NO = 1                                    '11/02/11
                    w利息未払前払.元金額 = G借入金テーブル(j).元金額                  '11/02/11
                    w利息未払前払.融資残高 = G借入金テーブル(j).融資残高        '11/02/11
                    
                    'うち入れ 未払い　利息計算対象額
                    If G借入金テーブル(j).据置x回目 = 1 Then                    '2012/03/01
                        w利息未払前払.利息計算対象額 = G借入金テーブル(j).元金額    '2012/03/01
                    Else                                                        '2012/03/01
                        w利息未払前払.利息計算対象額 = G借入金テーブル(j).融資残高 + G借入金テーブル(j).元金額  '11/02/16
                    End If                                                      '2012/03/01
                    
                    w利息未払前払.利息額増 = 0                                  '11/02/11
                    w利息未払前払.利息額減 = G借入金テーブル(j).利息額              '11/02/11
                    w未払利息残 = w未払利息残 - G借入金テーブル(j).利息額           '11/02/11
                    w利息未払前払.利息残高 = w未払利息残                        '11/02/11
                    w利息未払前払.日割日数 = G借入金テーブル(j).日割日数        '11/02/11
                    w利息未払前払.利率 = G借入金テーブル(j).利率                    '11/02/11
                    w利息未払前払.開始年月日 = DateAdd("d", 1 - G借入金テーブル(j).日割日数, w利息計算年月日) '11/02/11
                    w利息未払前払.終了年月日 = w利息計算年月日                  '11/02/11
                    
                    w利息未払前払.据置x回目 = G借入金テーブル(j).据置x回目
                    
                    
                    w利息未払前払.利息期間対象日数 = G借入金テーブル(j).日割日数    '2014/08/26
                    w利息未払前払.利息期間対象額 = G借入金テーブル(j).利息額        '2014/08/26
                    w利息未払前払.利息調整F = 0                                     '2014/08/26
                    
                    Call MBD010_利息未払前払Write(w利息未払前払, p借入計画マスタ)   '2016/09/26
                    
                    Call MRB010_未払利息増計算明細(p借入計画マスタ, _
                                               p借入計画マスタ.実行日, _
                                               w利息計算年月日, _
                                               G借入金テーブル(j).日割日数, _
                                               G借入金テーブル(j).利息額, _
                                               w解約実行日, _
                                               w解約締切年月日, _
                                               w基準年月)     '2016/09/26
                    
                End If                                              '2008/02/06 V182
                
               End If
               
             Next
             
             
             '***DCDA030=利息未払前払明細の作成
             Call MBD010_利息未払前払明細作成
             
             '***終了年月算出
             If Not IsNull(w解約実行日) Then                    '2008/02/07 V182
                w終了年月 = w解約実行日                         '2008/02/07 V182
             Else                                               '2008/02/07 V182
                w終了年月 = p借入計画マスタ.最終返済実行日      '2008/02/07 V182
             End If                                             '2008/02/07 V182
             
             w終了年月 = MBA010_対象年月(CDate(w終了年月))      '2008/02/07 V182
             
             w開始回目 = DateDiff("M", w基準年月, w基準年月) + 1 '2008/02/07 V182
             w終了回目 = DateDiff("M", w基準年月, w終了年月) + 1 '2008/02/07 V182
             
             '**************************************************************
             '     前払利息　未払利息　集計セット
             '**************************************************************
             For j = w開始回目 To w終了回目                     '2008/02/07 V182
                w対象年月 = DateAdd("M", j - 1, w基準年月)      '2008/02/07 V182
                If j = 1 Then                                   '2008/02/07 V182
                    p前払利息残(j) = p前払利息増(j) - p前払利息減(j) '2008/02/07V182
                    p未払利息残(j) = p未払利息増(j) - p未払利息減(j) '2008/02/07V182
                Else                                            '2008/02/07 V182
                    p前払利息残(j) = p前払利息残(j - 1) + p前払利息増(j) - p前払利息減(j) '2008/02/07 V182
                    p未払利息残(j) = p未払利息残(j - 1) + p未払利息増(j) - p未払利息減(j) '2008/02/07V182
                End If                                          '2008/02/07 V182
                
                For k = 1 To wcnt                               '2008/02/07 V182
                    If Format(w対象年月, "yyyymmdd") > Format(w年月(k - 1), "yyyymmdd") _
                        And Format(w対象年月, "yyyymmdd") <= Format(w年月(k), "yyyymmdd") Then '2008/02/07 V182
                            
                        w前払利息増(k) = w前払利息増(k) + p前払利息増(j) '2008/02/07 V182
                        w前払利息減(k) = w前払利息減(k) + p前払利息減(j) '2008/02/07 V182
                        w未払利息増(k) = w未払利息増(k) + p未払利息増(j) '2008/02/07 V182
                        w未払利息減(k) = w未払利息減(k) + p未払利息減(j) '2008/02/07 V182
                        
                        If Format(w対象年月, "yyyymmdd") = Format(w年月(k), "yyyymmdd") Then '2008/02/07 V182
                            w前払利息(k) = p前払利息残(j)       '2008/02/07 V182
                            w未払利息(k) = p未払利息残(j)       '2008/02/07 V182
                        End If                                  '2008/02/07 V182
                    End If                                      '2008/02/07 V182
                Next                                            '2008/02/07 V182
             Next                                               '2008/02/07 V182
             
             For k = 1 To wcnt
                If k > 1 Then
                    w前月残高(k) = w残高(k - 1)
                End If
                
                w残高(k) = w前月残高(k) + w融資(k) - w元金(k) - w解約(k)    ' 07/02/18 V180
                
             Next
             
          

               
             For k = 1 To wcnt
                If w解約(k) <> 0 Then               '11/06/17 V200
                    w元金(k) = w元金(k) + w解約(k)  '11/06/17 V200
                    w返済(k) = w返済(k) + w解約(k)  '11/06/17 V200
                End If                              '11/06/17 V200
                
                w融資合計 = w融資合計 + w融資(k)
                w元金合計 = w元金合計 + w元金(k)
                w利息合計 = w利息合計 + w利息(k)
                w返済合計 = w返済合計 + w返済(k)
                w解約合計 = w解約合計 + w解約(k)
                w保証合計 = w保証合計 + w保証(k)
                w手数料合計 = w手数料合計 + w手数料(k)                  '08/12/09 V189
                w前払利息増合計 = w前払利息増合計 + w前払利息増(k)      '2008/02/07 V182
                w前払利息減合計 = w前払利息減合計 + w前払利息減(k)      '2008/02/07 V182
                w未払利息増合計 = w未払利息増合計 + w未払利息増(k)      '2008/02/07 V182
                w未払利息減合計 = w未払利息減合計 + w未払利息減(k)      '2008/02/07 V182
                
             Next
             w残高合計 = w残高(wcnt)
             w前払利息合計 = w前払利息(wcnt)                            '2008/02/07 V182
             w未払利息合計 = w未払利息(wcnt)                            '2008/02/07 V182
                
             If w融資合計 = 0 And w元金合計 = 0 And w利息合計 = 0 And _
                w返済合計 = 0 And w解約合計 = 0 And w保証合計 = 0 And _
                w前払利息増合計 = 0 And w前払利息減合計 = 0 And _
                w未払利息増合計 = 0 And w未払利息減合計 = 0 And _
                w残高合計 = 0 Then
             Else
                 If FLG_Mdata = True Then
                    wRs2.AddNew
                        wRs2("借入番号") = w借入番号
                        wRs2("融資合計") = w融資合計
                        wRs2("元金合計") = w元金合計
                        wRs2("利息合計") = w利息合計
                        wRs2("返済合計") = w返済合計
                        wRs2("解約合計") = w解約合計
                        wRs2("保証合計") = w保証合計
                        wRs2("手数料合計") = w手数料合計            ' 08/12/09 V189
                        
                        wRs2("初期手数料合計") = w初期手数料合計
                        wRs2("元金手数料合計") = w元金手数料合計
                        wRs2("利息手数料合計") = w利息手数料合計
                        
                        
                        wRs2("残高合計") = w残高合計
    
                        wRs2("前払利息増合計") = w前払利息増合計    '2008/02/07 V182
                        wRs2("前払利息減合計") = w前払利息減合計    '2008/02/07 V182
                        wRs2("前払利息合計") = w前払利息合計        '2008/02/07 V182
                        wRs2("未払利息増合計") = w未払利息増合計    '2008/02/07 V182
                        wRs2("未払利息減合計") = w未払利息減合計    '2008/02/07 V182
                        wRs2("未払利息合計") = w未払利息合計        '2008/02/07 V182
    
    
                        For k = 1 To wcnt
                            wRs2("融資_" + CStr(Format(k, "00"))) = w融資(k)
                            wRs2("元金_" + CStr(Format(k, "00"))) = w元金(k)
                            wRs2("利息_" + CStr(Format(k, "00"))) = w利息(k)
                            wRs2("返済_" + CStr(Format(k, "00"))) = w返済(k)
                            wRs2("解約_" + CStr(Format(k, "00"))) = w解約(k)
                            wRs2("保証_" + CStr(Format(k, "00"))) = w保証(k)
                            wRs2("手数料_" + CStr(Format(k, "00"))) = w手数料(k)        ' 08/12/09 V189
                            
                            wRs2("初期手数料_" + CStr(Format(k, "00"))) = w初期手数料(k)
                            wRs2("元金手数料_" + CStr(Format(k, "00"))) = w元金手数料(k)
                            wRs2("利息手数料_" + CStr(Format(k, "00"))) = w利息手数料(k)
                            
                            wRs2("残高_" + CStr(Format(k, "00"))) = w残高(k)
    
                            wRs2("前払利息増_" + CStr(Format(k, "00"))) = w前払利息増(k) '2008/02/07 V182
                            wRs2("前払利息減_" + CStr(Format(k, "00"))) = w前払利息減(k) '2008/02/07 V182
                            wRs2("前払利息_" + CStr(Format(k, "00"))) = w前払利息(k)     '2008/02/07 V182
                            wRs2("未払利息増_" + CStr(Format(k, "00"))) = w未払利息増(k) '2008/02/07 V182
                            wRs2("未払利息減_" + CStr(Format(k, "00"))) = w未払利息減(k) '2008/02/07 V182
                            wRs2("未払利息_" + CStr(Format(k, "00"))) = w未払利息(k)     '2008/02/07 V182
                        Next
    
                    wRs2.Update
                
                 Else
                        
                    '2010/06/18
                    p推移(wiCnt).借入番号 = w借入番号
                    p推移(wiCnt).融資合計 = w融資合計
                    p推移(wiCnt).元金合計 = w元金合計
                    p推移(wiCnt).利息合計 = w利息合計
                    p推移(wiCnt).返済合計 = w返済合計
                    p推移(wiCnt).解約合計 = w解約合計
                    p推移(wiCnt).保証合計 = w保証合計
                    p推移(wiCnt).手数料合計 = w手数料合計
                    
                    p推移(wiCnt).初期手数料合計 = w初期手数料合計
                    p推移(wiCnt).元金手数料合計 = w元金手数料合計
                    p推移(wiCnt).利息手数料合計 = w利息手数料合計
                    
                    p推移(wiCnt).残高合計 = w残高合計

                    p推移(wiCnt).前払利息増合計 = w前払利息増合計
                    p推移(wiCnt).前払利息減合計 = w前払利息減合計
                    p推移(wiCnt).前払利息合計 = w前払利息合計
                    p推移(wiCnt).未払利息増合計 = w未払利息増合計
                    p推移(wiCnt).未払利息減合計 = w未払利息減合計
                    p推移(wiCnt).未払利息合計 = w未払利息合計

                    For k = 1 To wcnt
                        p推移(wiCnt).融資(k) = w融資(k)
                        p推移(wiCnt).元金(k) = w元金(k)
                        p推移(wiCnt).利息(k) = w利息(k)
                        p推移(wiCnt).返済(k) = w返済(k)
                        p推移(wiCnt).解約(k) = w解約(k)
                        p推移(wiCnt).保証(k) = w保証(k)
                        p推移(wiCnt).手数料(k) = w手数料(k)        ' 08/12/09 V189
                        
                        p推移(wiCnt).初期手数料(k) = w初期手数料(k)
                        p推移(wiCnt).元金手数料(k) = w元金手数料(k)
                        p推移(wiCnt).利息手数料(k) = w利息手数料(k)
                        
                        p推移(wiCnt).残高(k) = w残高(k)

                        p推移(wiCnt).前払利息増(k) = w前払利息増(k) '2008/02/07 V182
                        p推移(wiCnt).前払利息減(k) = w前払利息減(k) '2008/02/07 V182
                        p推移(wiCnt).前払利息(k) = w前払利息(k)     '2008/02/07 V182
                        p推移(wiCnt).未払利息増(k) = w未払利息増(k) '2008/02/07 V182
                        p推移(wiCnt).未払利息減(k) = w未払利息減(k) '2008/02/07 V182
                        p推移(wiCnt).未払利息(k) = w未払利息(k)     '2008/02/07 V182
                    Next

                    wiCnt = wiCnt + 1
                 End If
             End If
                
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing

    If FLG_Mdata = True Then
        wRs2.Close
        Set wRs = Nothing
    End If
'
    If FLG_Mdata <> True Then
        wstr2 = ""
        wstr2 = wstr2 + "Select * From DCDA010_借入残高推移表結果"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
        
            For wiCnt = 0 To 9999
            
                If p推移(wiCnt).借入番号 = "" Then
                    Exit For
                End If
                
                wRs2.AddNew
                    wRs2("借入番号") = p推移(wiCnt).借入番号
                    wRs2("融資合計") = p推移(wiCnt).融資合計
                    wRs2("元金合計") = p推移(wiCnt).元金合計
                    wRs2("利息合計") = p推移(wiCnt).利息合計
                    wRs2("返済合計") = p推移(wiCnt).返済合計
                    wRs2("解約合計") = p推移(wiCnt).解約合計
                    wRs2("保証合計") = p推移(wiCnt).保証合計
                    wRs2("手数料合計") = p推移(wiCnt).手数料合計
                    
                    wRs2("初期手数料合計") = p推移(wiCnt).初期手数料合計
                    wRs2("元金手数料合計") = p推移(wiCnt).元金手数料合計
                    wRs2("利息手数料合計") = p推移(wiCnt).利息手数料合計
                    
                    wRs2("残高合計") = p推移(wiCnt).残高合計
        
                    wRs2("前払利息増合計") = p推移(wiCnt).前払利息増合計
                    wRs2("前払利息減合計") = p推移(wiCnt).前払利息減合計
                    wRs2("前払利息合計") = p推移(wiCnt).前払利息合計
                    wRs2("未払利息増合計") = p推移(wiCnt).未払利息増合計
                    wRs2("未払利息減合計") = p推移(wiCnt).未払利息減合計
                    wRs2("未払利息合計") = p推移(wiCnt).未払利息合計
        
                    For k = 1 To wcnt
                        wRs2("融資_" + CStr(Format(k, "00"))) = p推移(wiCnt).融資(k)
                        wRs2("元金_" + CStr(Format(k, "00"))) = p推移(wiCnt).元金(k)
                        wRs2("利息_" + CStr(Format(k, "00"))) = p推移(wiCnt).利息(k)
                        wRs2("返済_" + CStr(Format(k, "00"))) = p推移(wiCnt).返済(k)
                        wRs2("解約_" + CStr(Format(k, "00"))) = p推移(wiCnt).解約(k)
                        wRs2("保証_" + CStr(Format(k, "00"))) = p推移(wiCnt).保証(k)
                        wRs2("手数料_" + CStr(Format(k, "00"))) = p推移(wiCnt).手数料(k)
                        
                        wRs2("初期手数料_" + CStr(Format(k, "00"))) = p推移(wiCnt).初期手数料(k)
                        wRs2("元金手数料_" + CStr(Format(k, "00"))) = p推移(wiCnt).元金手数料(k)
                        wRs2("利息手数料_" + CStr(Format(k, "00"))) = p推移(wiCnt).利息手数料(k)
                        
                        wRs2("残高_" + CStr(Format(k, "00"))) = p推移(wiCnt).残高(k)
        
                        wRs2("前払利息増_" + CStr(Format(k, "00"))) = p推移(wiCnt).前払利息増(k)
                        wRs2("前払利息減_" + CStr(Format(k, "00"))) = p推移(wiCnt).前払利息減(k)
                        wRs2("前払利息_" + CStr(Format(k, "00"))) = p推移(wiCnt).前払利息(k)
                        wRs2("未払利息増_" + CStr(Format(k, "00"))) = p推移(wiCnt).未払利息増(k)
                        wRs2("未払利息減_" + CStr(Format(k, "00"))) = p推移(wiCnt).未払利息減(k)
                        wRs2("未払利息_" + CStr(Format(k, "00"))) = p推移(wiCnt).未払利息(k)
                    Next
        
                wRs2.Update
        
            Next wiCnt
            
        wRs2.Close
        Set wRs2 = Nothing
    End If
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MRB010_標準入力未払前払_ERR:
    pERR_MES = pPROGRAM_ID + "/ MRB010_標準入力未払前払() でエラー" + vbCrLf + vbCrLf + _
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
' MRB010_前払利息減計算
'------------------------------------------------
Public Sub MRB010_前払利息減計算(p借入計画 As MAA910_借入金, _
                                 p実行年月日 As Variant, _
                                 p実際年月日 As Variant, _
                                 p利息対象期間日数 As Integer, _
                                 p利息額 As Double, _
                                 p解約年月日 As Variant, _
                                 p解約締切年月日 As Variant, _
                                 p基準年月 As Date)
'
    Dim w前払利息減 As Double
    Dim w前払利息減計 As Double
    Dim w前月前払利息残 As Double
    Dim w対象年月 As Date
    Dim w開始年月日 As Date
    Dim w最終年月日 As Date
    Dim w締切年月日 As Date
    Dim w計算対象日数 As Integer
    Dim w回目 As Integer
    Dim p回目 As Integer
'
    On Error GoTo MRB010_前払利息減計算_ERR
    
    '***借入番号確認
    p借入計画.借入番号 = p借入計画.借入番号         '2016/09/27
'
    '***内入で日割日数が、マイナス、になった時　日数を絶対値に変換
    'If Format(p実際年月日, "yyyy/mm/dd") <> Format(p解約年月日, "yyyy/mm/dd") Then  ' 10/01/09
        If p利息対象期間日数 < 0 Then                                               ' 08/12/17 V189
            p利息対象期間日数 = p利息対象期間日数 * -1                              ' 08/12/17 V189
        End If                                                                      ' 08/12/17 V189
    'End If                                                                          ' 10/01/09
    
    w前払利息減計 = 0
    'If Format(p実際年月日, "yyyymmdd") <> Format(p実行年月日, "yyyymmdd") Then
    '    w開始年月日 = DateAdd("D", 1, p実際年月日)
    '    w最終年月日 = DateAdd("D", p利息対象期間日数, p実際年月日)
    'Else
    '    w開始年月日 = p実行年月日
    '    w最終年月日 = DateAdd("D", -1 + p利息対象期間日数, p実行年月日)
    'End If
    
    w開始年月日 = p実際年月日
    w最終年月日 = DateAdd("D", -1 + p利息対象期間日数, p実際年月日)
    
    
    
    
    
    
    '***中間利払いの時　開始年月日　終了年月日　を　1日減ずる
    If p借入計画.利息控除区分 = 4 Then              '2016/09/27
        If p借入計画.利息区分 = XMXA020_区分("利息区分", "利息先払") Then               '2016/09/28
            '**利息先払いの時
            If Format(p借入計画.実行日, "yyyy/mm/dd") = Format(p実際年月日, "yyyy/mm/dd") Then  '2016/09/27
            Else                                    '2016/09/27
                w開始年月日 = DateAdd("d", -1, w開始年月日)     '2016/09/27
                w最終年月日 = DateAdd("d", -1, w最終年月日)     '2016/09/27
            End If                                              '2016/09/27
        Else                                                    '2016/09/27
            '**利息後払いの時
            w開始年月日 = DateAdd("d", -1, w開始年月日)     '2016/09/27
            w最終年月日 = DateAdd("d", -1, w最終年月日)     '2016/09/27
        End If                                              '2016/09/27
    End If
    
    
    
            
    
    
    If Format(p解約年月日, "yyyymmdd") <> Format(p実際年月日, "yyyymmdd") Then
STA1:
        w締切年月日 = MBA010_締日年月日(CDate(w開始年月日))
        
        If Format(p解約締切年月日, "yyyymmdd") = Format(w締切年月日, "yyyymmdd") Then
            Exit Sub
        Else
            w計算対象日数 = DateDiff("D", w開始年月日, w締切年月日) + 1
        End If
        
        w対象年月 = MBA010_対象年月(CDate(w開始年月日))
        w回目 = DateDiff("M", p基準年月, w対象年月) + 1
        
        
        If Format(w最終年月日, "yyyymmdd") > Format(w締切年月日, "yyyymmdd") Then
            w前払利息減 = Fix(p利息額 * w計算対象日数 / p利息対象期間日数) '09/12/25
            'w前払利息減 = (w前払利息減 + 5)                        '09/12/25
            'w前払利息減 = Fix(w前払利息減 / 10)                     '09/12/25
            If w回目 > 0 And w回目 <= 600 Then
                p前払利息減(w回目) = p前払利息減(w回目) + w前払利息減
            End If
            
            w前払利息減計 = w前払利息減計 + w前払利息減
            w開始年月日 = DateAdd("D", 1, w締切年月日)
        Else
            If w回目 > 0 And w回目 <= 600 Then
                p前払利息減(w回目) = p前払利息減(w回目) + (p利息額 - w前払利息減計)
            End If
            
            Exit Sub
        End If
        
        GoTo STA1
        
    Else
    '**解約
        w対象年月 = MBA010_対象年月(CDate(p解約年月日))
        w回目 = DateDiff("M", p基準年月, w対象年月) + 1
        
        For p回目 = 1 To w回目 - 1
            If p回目 = 1 Then
                If w回目 > 0 And w回目 <= 600 Then
                    w前月前払利息残 = p前払利息増(p回目) - p前払利息減(p回目)
                End If
                
            Else
                w前月前払利息残 = w前月前払利息残 + p前払利息増(p回目) - p前払利息減(p回目)
            End If
        Next
        
        If w回目 > 0 And w回目 <= 600 Then
            'p前払利息減(w回目) = p前払利息減(w回目) + w前月前払利息残 + p前払利息増(w回目)
            p前払利息減(w回目) = w前月前払利息残 + p前払利息増(w回目)           '10/02/13
            
        End If
        
        
        Exit Sub
    End If
'
'----------< ERROR ROUTINE >---------------------------------------------------
MRB010_前払利息減計算_ERR:
    pERR_MES = pPROGRAM_ID + "/ MRB010_前払利息減計算() でエラー" + vbCrLf + vbCrLf + _
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
' MRB010_前払利息減計算明細　　11/02/11
'------------------------------------------------
Public Sub MRB010_前払利息減計算明細(p借入計画 As MAA910_借入金, _
                                 p実行年月日 As Variant, _
                                 p利息控除区分 As Integer, _
                                 p実際年月日 As Variant, _
                                 p利息対象期間日数 As Integer, _
                                 p利息額 As Double, _
                                 p解約年月日 As Variant, _
                                 p解約締切年月日 As Variant, _
                                 p基準年月 As Date)
'
    Dim w前払利息減 As Double
    Dim w前払利息減計 As Double
    Dim w前月前払利息残 As Double
    Dim w対象年月 As Date
    Dim w開始年月日 As Date
    Dim w最終年月日 As Date
    Dim w締切年月日 As Date
    Dim w計算対象日数 As Integer
    Dim w回目 As Integer
    Dim p回目 As Integer
    Dim w当初日数 As Integer
    
'
    w当初日数 = p利息対象期間日数
    On Error GoTo MRB010_前払利息減計算明細_ERR
    
    
'
    '***内入で日割日数が、マイナス、になった時　日数を絶対値に変換
    'If Format(p実際年月日, "yyyy/mm/dd") <> Format(p解約年月日, "yyyy/mm/dd") Then  ' 10/01/09
        If p利息対象期間日数 < 0 Then                                               ' 08/12/17 V189
            p利息対象期間日数 = p利息対象期間日数 * -1                              ' 08/12/17 V189
        End If                                                                      ' 08/12/17 V189
    'End If                                                                          ' 10/01/09
    
    w前払利息減計 = 0
    'If Format(p実際年月日, "yyyymmdd") <> Format(p実行年月日, "yyyymmdd") Then
    '    w開始年月日 = DateAdd("D", 1, p実際年月日)
    '    w最終年月日 = DateAdd("D", p利息対象期間日数, p実際年月日)
    'Else
    '    w開始年月日 = p実行年月日
    '    w最終年月日 = DateAdd("D", -1 + p利息対象期間日数, p実行年月日)
    'End If
    
    w開始年月日 = p実際年月日
    w最終年月日 = DateAdd("D", -1 + p利息対象期間日数, p実際年月日)
    
    
    
    '***中間利払いの時　開始年月日　終了年月日　を　1日減ずる
    If p借入計画.利息控除区分 = 4 Then              '2016/09/27
        If p借入計画.利息区分 = XMXA020_区分("利息区分", "利息先払") Then               '2016/09/28
            '**利息先払いの時
            If Format(p借入計画.実行日, "yyyy/mm/dd") = Format(p実際年月日, "yyyy/mm/dd") Then  '2016/09/27
            Else                                    '2016/09/27
                w開始年月日 = DateAdd("d", -1, w開始年月日)     '2016/09/27
                w最終年月日 = DateAdd("d", -1, w最終年月日)     '2016/09/27
            End If                                              '2016/09/27
        Else                                                    '2016/09/27
            '**利息後払いの時
            w開始年月日 = DateAdd("d", -1, w開始年月日)     '2016/09/27
            w最終年月日 = DateAdd("d", -1, w最終年月日)     '2016/09/27
        End If                                              '2016/09/27
    End If

    
    
    
    
    
    If Format(p解約年月日, "yyyymmdd") <> Format(p実際年月日, "yyyymmdd") Then
STA1:
        w締切年月日 = MBA010_締日年月日(CDate(w開始年月日))
        
        If Format(p解約締切年月日, "yyyymmdd") = Format(w締切年月日, "yyyymmdd") Then
            
            '***その2　（最終回の時）
            'w利息未払前払.返済年月日 = w締切年月日
            w利息未払前払.月毎NO = 0
            w利息未払前払.元金額 = 0
            w利息未払前払.利息額増 = 0
            w利息未払前払.利息額減 = p利息額 - w前払利息減計
            w利息未払前払.日割日数 = Round(w利息未払前払.利息額減 * 365 * 100 / w利息未払前払.利息計算対象額 / w利息未払前払.利率)
            
            w利息未払前払.開始年月日 = w開始年月日
            
            If w利息未払前払.日割日数 < 0 Then              '2012/03/06
                w利息未払前払.返済年月日 = DateAdd("d", -w利息未払前払.日割日数 - 1, w利息未払前払.開始年月日)
            Else                                            '2012/03/06
                w利息未払前払.返済年月日 = DateAdd("d", w利息未払前払.日割日数 - 1, w利息未払前払.開始年月日)
            End If                                          '2012/03/06
            
            w利息未払前払.終了年月日 = w利息未払前払.返済年月日
            
            If w利息未払前払.返済年月日 <= p解約年月日 Then
                w前払利息残 = w前払利息残 - (p利息額 - w前払利息減計)
                w利息未払前払.利息残高 = w前払利息残
                
                w利息未払前払.利息期間対象日数 = p利息対象期間日数  '2014/08/26
                w利息未払前払.利息期間対象額 = p利息額              '2014/08/26
                w利息未払前払.利息調整F = 1                         '2014/08/26
                
                Call MBD010_利息未払前払Write(w利息未払前払, p借入計画)     '2016/09/26
            End If
            
            Exit Sub
        Else
            w計算対象日数 = DateDiff("D", w開始年月日, w締切年月日) + 1
        End If
        
        w対象年月 = MBA010_対象年月(CDate(w開始年月日))
        w回目 = DateDiff("M", p基準年月, w対象年月) + 1
        
        
        If Format(w最終年月日, "yyyymmdd") > Format(w締切年月日, "yyyymmdd") Then
            w前払利息減 = Fix(p利息額 * w計算対象日数 / p利息対象期間日数) '09/12/25
            'w前払利息減 = (w前払利息減 + 5)                        '09/12/25
            'w前払利息減 = Fix(w前払利息減 / 10)                     '09/12/25
            If w回目 > 0 And w回目 <= 600 Then
                p前払利息減(w回目) = p前払利息減(w回目) + w前払利息減
            End If
            
            '***その１　（最終回以前の時）
            w利息未払前払.返済年月日 = w締切年月日
            w利息未払前払.月毎NO = 0
            w利息未払前払.元金額 = 0
            w利息未払前払.利息額増 = 0
            w利息未払前払.利息額減 = w前払利息減
            w利息未払前払.日割日数 = w計算対象日数
            w利息未払前払.開始年月日 = w開始年月日
            w利息未払前払.終了年月日 = w締切年月日
            w前払利息残 = w前払利息残 - w前払利息減
            w利息未払前払.利息残高 = w前払利息残
            w当初日数 = w当初日数 - w利息未払前払.日割日数  '11/02/16
            
            w利息未払前払.利息期間対象日数 = p利息対象期間日数  '2014/08/26
            w利息未払前払.利息期間対象額 = p利息額              '2014/08/26
            w利息未払前払.利息調整F = 0                         '2014/08/26
            
            Call MBD010_利息未払前払Write(w利息未払前払, p借入計画)     '2016/09/26
            
            w前払利息減計 = w前払利息減計 + w前払利息減
            w開始年月日 = DateAdd("D", 1, w締切年月日)
        Else
            If w回目 > 0 And w回目 <= 600 Then
                p前払利息減(w回目) = p前払利息減(w回目) + (p利息額 - w前払利息減計)
            End If
            
            '***その2　（最終回の時）
            'w利息未払前払.返済年月日 = w締切年月日
            w利息未払前払.月毎NO = 0
            w利息未払前払.元金額 = 0
            w利息未払前払.利息額増 = 0
            w利息未払前払.利息額減 = p利息額 - w前払利息減計
            
            If w手入力区分 = 0 Then
                w利息未払前払.日割日数 = Round(w利息未払前払.利息額減 * 365 * 100 / w利息未払前払.利息計算対象額 / w利息未払前払.利率)
            Else
                w利息未払前払.日割日数 = w当初日数
            End If
            
            w利息未払前払.開始年月日 = w開始年月日
            
            If w利息未払前払.日割日数 < 0 Then          '2012/03/06
                w利息未払前払.返済年月日 = DateAdd("d", -w利息未払前払.日割日数 - 1, w利息未払前払.開始年月日)
            Else                                        '2012/03/06
                w利息未払前払.返済年月日 = DateAdd("d", w利息未払前払.日割日数 - 1, w利息未払前払.開始年月日)
            End If                                      '2012/03/06
            
            w利息未払前払.終了年月日 = w利息未払前払.返済年月日
            w前払利息残 = w前払利息残 - (p利息額 - w前払利息減計)
            w利息未払前払.利息残高 = w前払利息残
            
            w利息未払前払.利息期間対象日数 = p利息対象期間日数  '2014/08/26
            w利息未払前払.利息期間対象額 = p利息額              '2014/08/26
            w利息未払前払.利息調整F = 1                         '2014/08/26
            
            Call MBD010_利息未払前払Write(w利息未払前払, p借入計画)
            
            Exit Sub
        End If
        
        GoTo STA1
        
    Else
    '**解約
        w対象年月 = MBA010_対象年月(CDate(p解約年月日))
        w回目 = DateDiff("M", p基準年月, w対象年月) + 1
        
        For p回目 = 1 To w回目 - 1
            If p回目 = 1 Then
                If w回目 > 0 And w回目 <= 600 Then
                    w前月前払利息残 = p前払利息増(p回目) - p前払利息減(p回目)
                End If
                
            Else
                w前月前払利息残 = w前月前払利息残 + p前払利息増(p回目) - p前払利息減(p回目)
            End If
        Next
        
        If w回目 > 0 And w回目 <= 600 Then
            'p前払利息減(w回目) = p前払利息減(w回目) + w前月前払利息残 + p前払利息増(w回目)
            p前払利息減(w回目) = w前月前払利息残 + p前払利息増(w回目)           '10/02/13
            
        End If
        
        '***その３（解約の時）
        w利息未払前払.返済年月日 = p解約年月日                          '11/02/15
        w利息未払前払.月毎NO = 0                                        '11/02/15
        w利息未払前払.元金額 = w利息未払前払.融資残高                   '11/02/15
        w利息未払前払.利息額増 = 0                                      '11/02/15
        'w利息未払前払.利息額減 = w利息未払前払.利息残高                 '12/07/12
        w利息未払前払.利息額減 = p前払利息減(w回目)                     '12/07/12
        
        
        If p利息控除区分 >= 2 Then                                      '2012/03/23
            w利息未払前払.終了年月日 = DateAdd("d", -1, p解約年月日)    '2012/03/23
        Else                                                            '2012/03/23
            w利息未払前払.終了年月日 = p解約年月日                          '11/02/15
        End If                                                          '2012/03/23
        
        w利息未払前払.日割日数 = Round(w利息未払前払.利息額減 * 365 * 100 _
                                       / w利息未払前払.利息計算対象額 / w利息未払前払.利率) '11/02/15
                                       
        w利息未払前払.開始年月日 = DateAdd("d", -w利息未払前払.日割日数 + 1, w利息未払前払.終了年月日) '11/02/15
        w利息未払前払.融資残高 = 0                                      '11/02/15
        w利息未払前払.利息残高 = 0                                      '11/02/15
        
        'If w利息未払前払.利息額減 <> 0 Then                               '2012/07/10
        
        w利息未払前払.利息調整F = 2                                     '2014/08/26
        
            Call MBD010_利息未払前払Write(w利息未払前払, p借入計画)     '2016/09/26
        'End If                                                          '2012/07/10
        
        
        Exit Sub
    End If
'
'----------< ERROR ROUTINE >---------------------------------------------------
MRB010_前払利息減計算明細_ERR:
    pERR_MES = pPROGRAM_ID + "/ MRB010_前払利息減計算明細() でエラー" + vbCrLf + vbCrLf + _
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
' MRB010_未払利息増計算
'------------------------------------------------
Public Sub MRB010_未払利息増計算(p借入計画 As MAA910_借入金, _
                                 p実行年月日 As Variant, _
                                 p実際年月日 As Variant, _
                                 p利息対象期間日数 As Integer, _
                                 p利息額 As Double, _
                                 p解約年月日 As Variant, _
                                 p解約締切年月日 As Variant, _
                                 p基準年月 As Date)
'
    Dim w未払利息増 As Double
    Dim w未払利息増計 As Double
    Dim w前月未払利息残 As Double
    Dim w対象年月 As Date
    Dim w開始年月日 As Date
    Dim w最終年月日 As Date
    Dim w締切年月日 As Date
    Dim w計算対象日数 As Integer
    Dim w回目 As Integer
    Dim p回目 As Integer
'
    On Error GoTo MRB010_未払利息増計算_ERR
    
    '***借入番号確認
    p借入計画.借入番号 = p借入計画.借入番号         '2016/09/27
'
    w未払利息増計 = 0
    
    w開始年月日 = DateAdd("D", 1 - p利息対象期間日数, p実際年月日)
    
    w最終年月日 = p実際年月日
    
    
    '***中間利払いの時　開始年月日　終了年月日　を　1日減ずる
    If p借入計画.利息控除区分 = 4 Then              '2016/09/27
        If p借入計画.利息区分 = XMXA020_区分("利息区分", "利息先払") Then               '2016/09/28
            '**利息先払いの時
            If Format(p借入計画.実行日, "yyyy/mm/dd") = Format(p実際年月日, "yyyy/mm/dd") Then  '2016/09/27
            Else                                    '2016/09/27
                w開始年月日 = DateAdd("d", -1, w開始年月日)     '2016/09/27
                w最終年月日 = DateAdd("d", -1, w最終年月日)     '2016/09/27
            End If                                              '2016/09/27
        Else                                                    '2016/09/27
            '**利息後払いの時
            w開始年月日 = DateAdd("d", -1, w開始年月日)     '2016/09/27
            w最終年月日 = DateAdd("d", -1, w最終年月日)     '2016/09/27
        End If                                              '2016/09/27
    End If

    
    
    
    
STA1:
    w締切年月日 = MBA010_締日年月日(CDate(w開始年月日))
    
    w計算対象日数 = DateDiff("D", w開始年月日, w締切年月日) + 1
       
        
    w対象年月 = MBA010_対象年月(CDate(w開始年月日))
    w回目 = DateDiff("M", p基準年月, w対象年月) + 1
    
              
    
        
        
    If Format(w最終年月日, "yyyymmdd") > Format(w締切年月日, "yyyymmdd") Then
        w未払利息増 = Fix(p利息額 * w計算対象日数 / p利息対象期間日数)  '09/12/25
        'w未払利息増 = (w未払利息増 + 5)                           '09/12/25
        'w未払利息増 = Fix(w未払利息増 / 10)                         '09/12/25
        
        If w回目 > 0 And w回目 <= 600 Then
            p未払利息増(w回目) = p未払利息増(w回目) + w未払利息増
        End If
        
        w未払利息増計 = w未払利息増計 + w未払利息増
        
        w開始年月日 = DateAdd("D", 1, w締切年月日)
         
        GoTo STA1
        
    Else
        If w回目 > 0 And w回目 <= 600 Then
            p未払利息増(w回目) = p未払利息増(w回目) + p利息額 - w未払利息増計
        End If
        
        Exit Sub
    
    End If
'
'----------< ERROR ROUTINE >---------------------------------------------------
MRB010_未払利息増計算_ERR:
    pERR_MES = pPROGRAM_ID + "/ MRB010_未払利息増計算() でエラー" + vbCrLf + vbCrLf + _
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
' MRB010_未払利息増計算明細
'------------------------------------------------
Public Sub MRB010_未払利息増計算明細(p借入計画 As MAA910_借入金, _
                                 p実行年月日 As Variant, _
                                 p実際年月日 As Variant, _
                                 p利息対象期間日数 As Integer, _
                                 p利息額 As Double, _
                                 p解約年月日 As Variant, _
                                 p解約締切年月日 As Variant, _
                                 p基準年月 As Date)
'
    Dim w未払利息増 As Double
    Dim w未払利息増計 As Double
    Dim w前月未払利息残 As Double
    Dim w対象年月 As Date
    Dim w開始年月日 As Date
    Dim w最終年月日 As Date
    Dim w締切年月日 As Date
    Dim w計算対象日数 As Integer
    Dim w回目 As Integer
    Dim p回目 As Integer
    Dim w当初日数 As Integer
    
'
    On Error GoTo MRB010_未払利息増計算明細_ERR
    
    'If w利息未払前払.据置x回目 = 1 Then     '2012/03/01
    '    w利息未払前払.利息計算対象額 = w利息未払前払.元金額     '2012/03/01
    'End If                                                      '2012/03/01
    
    
    w当初日数 = p利息対象期間日数
    
    GoTo w無視              '2012/03/08
    
    
    If p解約年月日 = p実際年月日 Then                   '11/02/15
         '***その３（解約の時）
        w利息未払前払.返済年月日 = p解約年月日                          '11/02/15
        w利息未払前払.月毎NO = 0                                        '11/02/15
        w利息未払前払.元金額 = w利息未払前払.融資残高                   '11/02/15
        w利息未払前払.利息額減 = 0                                      '11/02/15
        w利息未払前払.利息額増 = -w利息未払前払.利息残高                 '11/02/15
        w利息未払前払.終了年月日 = p解約年月日                          '11/02/15
        w利息未払前払.日割日数 = Round(w利息未払前払.利息額増 * 365 * 100 _
                                       / w利息未払前払.利息計算対象額 / w利息未払前払.利率) '11/02/15
                                       
        w利息未払前払.開始年月日 = DateAdd("d", -w利息未払前払.日割日数 + 1, w利息未払前払.終了年月日) '11/02/15
        w利息未払前払.融資残高 = 0                                      '11/02/15
        w利息未払前払.利息残高 = 0                                      '11/02/15
        
        Call MBD010_利息未払前払Write(w利息未払前払, p借入計画)         '2016/09/26
        
        
        Exit Sub
    End If
    
w無視:                  '2012/03/08


    
'
    w未払利息増計 = 0
    
    w開始年月日 = DateAdd("D", 1 - p利息対象期間日数, p実際年月日)
    
    w最終年月日 = p実際年月日
    
    
    
    '***中間利払いの時　開始年月日　終了年月日　を　1日減ずる
    If p借入計画.利息控除区分 = 4 Then              '2016/09/27
        If p借入計画.利息区分 = XMXA020_区分("利息区分", "利息先払") Then               '2016/09/28
            '**利息先払いの時
            If Format(p借入計画.実行日, "yyyy/mm/dd") = Format(p実際年月日, "yyyy/mm/dd") Then  '2016/09/27
            Else                                    '2016/09/27
                w開始年月日 = DateAdd("d", -1, w開始年月日)     '2016/09/27
                w最終年月日 = DateAdd("d", -1, w最終年月日)     '2016/09/27
            End If                                              '2016/09/27
        Else                                                    '2016/09/27
            '**利息後払いの時
            w開始年月日 = DateAdd("d", -1, w開始年月日)     '2016/09/27
            w最終年月日 = DateAdd("d", -1, w最終年月日)     '2016/09/27
        End If                                              '2016/09/27
    End If

    
    
    
 
STA1:
    w締切年月日 = MBA010_締日年月日(CDate(w開始年月日))
    
    w計算対象日数 = DateDiff("D", w開始年月日, w締切年月日) + 1
       
        
    w対象年月 = MBA010_対象年月(CDate(w開始年月日))
    w回目 = DateDiff("M", p基準年月, w対象年月) + 1
        
        
    If Format(w最終年月日, "yyyymmdd") > Format(w締切年月日, "yyyymmdd") Then
        w未払利息増 = Fix(p利息額 * w計算対象日数 / p利息対象期間日数)  '09/12/25
        'w未払利息増 = (w未払利息増 + 5)                           '09/12/25
        'w未払利息増 = Fix(w未払利息増 / 10)                         '09/12/25
        
        If w回目 > 0 And w回目 <= 600 Then
            p未払利息増(w回目) = p未払利息増(w回目) + w未払利息増
        End If
        
        '***その１　（最終回以前の時）
        w利息未払前払.返済年月日 = w締切年月日
        w利息未払前払.月毎NO = 0
        w利息未払前払.元金額 = 0
        w利息未払前払.利息額増 = w未払利息増
        w利息未払前払.利息額減 = 0
        w利息未払前払.日割日数 = w計算対象日数
        w利息未払前払.開始年月日 = w開始年月日
        w利息未払前払.終了年月日 = w締切年月日
        w未払利息残 = w未払利息残 + w未払利息増
        w利息未払前払.利息残高 = w未払利息残
        w当初日数 = w当初日数 - w利息未払前払.日割日数      '11/02/16
        w利息未払前払.利息期間対象日数 = p利息対象期間日数  '2014/08/26
        w利息未払前払.利息期間対象額 = p利息額              '2014/08/26
        w利息未払前払.利息調整F = 0                         '2014/08/26
        
        Call MBD010_利息未払前払Write(w利息未払前払, p借入計画)     '2016/09/26
        
        
        w未払利息増計 = w未払利息増計 + w未払利息増
        
        w開始年月日 = DateAdd("D", 1, w締切年月日)
         
        GoTo STA1
        
    Else
        If w回目 > 0 And w回目 <= 600 Then
            p未払利息増(w回目) = p未払利息増(w回目) + p利息額 - w未払利息増計
        End If
        
        '***その2　（最終回時）
        'w利息未払前払.返済年月日 = w締切年月日
        w利息未払前払.月毎NO = 0
        w利息未払前払.元金額 = 0
        w利息未払前払.利息額増 = p利息額 - w未払利息増計
        w利息未払前払.利息額減 = 0
        
        If w手入力区分 = 0 Then
            w利息未払前払.日割日数 = Round(w利息未払前払.利息額増 * 365 * 100 / w利息未払前払.利息計算対象額 / w利息未払前払.利率)
        Else
            w利息未払前払.日割日数 = w当初日数                 '11/02/16
        End If
        
            
        w利息未払前払.開始年月日 = w開始年月日
        w利息未払前払.終了年月日 = DateAdd("d", w利息未払前払.日割日数 - 1, w利息未払前払.開始年月日)
        w利息未払前払.返済年月日 = w利息未払前払.終了年月日
        w未払利息残 = w未払利息残 + p利息額 - w未払利息増計
        w利息未払前払.利息残高 = w未払利息残
        
        w利息未払前払.利息期間対象日数 = p利息対象期間日数  '2014/08/26
        w利息未払前払.利息期間対象額 = p利息額              '2014/08/26
        w利息未払前払.利息調整F = 1                         '2014/08/26
        
        Call MBD010_利息未払前払Write(w利息未払前払, p借入計画)     '2016/09/26
        
        Exit Sub
    
    End If
'
'----------< ERROR ROUTINE >---------------------------------------------------
MRB010_未払利息増計算明細_ERR:
    pERR_MES = pPROGRAM_ID + "/ MRB010_未払利息増計算明細() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub

'--------------------------------------------------------
' MBA010_支払年月日算出   08/03/05 V185
'--------------------------------------------------------
Public Function MBA010_支払年月日算出(p対象年月日 As Date, p支払日 As Integer) As Date
'
    Dim w対象年月日 As Date
    Dim w年 As Integer
    Dim w月 As Integer
    Dim w日 As Integer
    Dim w閏年 As Integer
    
'
    On Error GoTo MBA010_支払年月日算出_ERR
    
'
    w対象年月日 = p対象年月日               ' 07/03/03 V180
    w年 = Year(w対象年月日)
    w月 = Month(w対象年月日)
    
    w日 = Day(w対象年月日)
    If w日 > p支払日 Then
        w対象年月日 = DateAdd("m", 1, w対象年月日)
    End If
    
    
    w年 = Year(w対象年月日)
    w月 = Month(w対象年月日)
    
    If 31 = p支払日 Then
        MBA010_支払年月日算出 = CDate(CStr(w年) + "/" + CStr(w月) + "/" + CStr(1))
        MBA010_支払年月日算出 = DateAdd("M", 1, MBA010_支払年月日算出)
        MBA010_支払年月日算出 = DateDiff("D", 1, MBA010_支払年月日算出)
    Else
        
        If w月 = 2 And p支払日 > 28 Then
            w閏年 = w年 Mod 4
            If w閏年 = 0 Then
                MBA010_支払年月日算出 = CDate(CStr(w年) + "/" + CStr(w月) + "/" + CStr(29))
            Else
                MBA010_支払年月日算出 = CDate(CStr(w年) + "/" + CStr(w月) + "/" + CStr(28))
            End If
        Else
            MBA010_支払年月日算出 = CDate(CStr(w年) + "/" + CStr(w月) + "/" + CStr(p支払日))
        End If
        
    End If
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBA010_支払年月日算出_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBA010_支払年月日算出() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
    
End Function

'------------------------------------------------
' MBD010_利息未払前払Write 2011/02/11
'------------------------------------------------
Private Sub MBD010_利息未払前払Write(p利息未払前払 As MAA910_利息未払前払テーブル, p借入計画 As MAA910_借入金)  '2016/09/26
'
    Dim wFind As Boolean
    Dim w配列数 As Integer
    Dim j As Integer
'
    On Error GoTo MBD010_利息未払前払Write_ERR
    
    'If p利息未払前払.利息区分 = "1" _
    '    And p利息未払前払.利息額増 = 0 _
    '    And p利息未払前払.利息額減 = 0 Then         '2012/07/12
    '    Exit Sub                                    '2012/07/12
    'End If                                          '2012/07/11
    
    
  'If p借入金テーブル.元金額 <> 0 Or p借入金テーブル.利息額 <> 0 Or p借入金テーブル.保証料 <> 0 _
  '   Or p借入金テーブル.手数料 Then         '10/02/27
'
    w配列数 = UBound(G利息未払前払テーブル)
    
    wFind = False           '2012/07/11
    For j = 1 To w配列数    '2012/07/11
      If p利息未払前払.利息区分 = "1" _
            And G利息未払前払テーブル(j).借入番号 = p利息未払前払.借入番号 _
            And G利息未払前払テーブル(j).返済年月日 = p利息未払前払.返済年月日 _
            And G利息未払前払テーブル(j).利息額増 = 0 _
            And p利息未払前払.利息額増 = 0 Then                                     '2012/07/11
                
            wFind = True                                                            '2012/07/11
            Exit For                                                                '2012/07/11
      
      End If                                                                      '2012/07/11
        
    Next                                                                            '2012/07/11
    
            
            
            
            

    
    'For j = 1 To w配列数
    '    If p借入金テーブル.借入番号 = G借入金テーブル(j).借入番号 _
    '    And p借入金テーブル.実際年月日 = G借入金テーブル(j).実際年月日 _
    '    And p借入金テーブル.据置X回目 = G借入金テーブル(j).据置X回目 Then     ' 08/12/06 V189

    '        wFind = True
    '        Exit For
    '    End If
    'Next
    
    '   And p借入金テーブル.返済回数 = G借入金テーブル(J).返済回数 _
    '    And p借入金テーブル.据置X回目 = G借入金テーブル(J).据置X回目 _

    If p利息未払前払.利息区分 = "1" _
        And p利息未払前払.据置x回目 = 3 Then            '2012/07/12
        wFind = False                                   '2012/07/12
    End If                                              '2012/07/12
    
        

    If Not wFind Then
        j = w配列数 + 1
        ReDim Preserve G利息未払前払テーブル(j)
    End If

    '** テーブルにセット **
    G利息未払前払テーブル(j).借入番号 = p利息未払前払.借入番号
    G利息未払前払テーブル(j).銀行番号 = p利息未払前払.銀行番号
    G利息未払前払テーブル(j).利息区分 = p利息未払前払.利息区分
    G利息未払前払テーブル(j).利息計算日数区分 = p利息未払前払.利息計算日数区分
    G利息未払前払テーブル(j).返済年月日 = p利息未払前払.返済年月日
    G利息未払前払テーブル(j).月毎NO = p利息未払前払.月毎NO
    G利息未払前払テーブル(j).元金額 = p利息未払前払.元金額
    G利息未払前払テーブル(j).融資残高 = p利息未払前払.融資残高
    
    'If p利息未払前払.据置x回目 = 1 Or p利息未払前払.据置x回目 = 3 Then      '2012/03/06
    '    G利息未払前払テーブル(j).利息計算対象額 = p利息未払前払.元金額      '2012/03/06
    'Else                                                                    '2012/03/06
    G利息未払前払テーブル(j).利息計算対象額 = p利息未払前払.利息計算対象額
    'End If                                                                  '2012/03/06
    
    G利息未払前払テーブル(j).利息額増 = p利息未払前払.利息額増
    G利息未払前払テーブル(j).利息額減 = p利息未払前払.利息額減
    G利息未払前払テーブル(j).利息残高 = p利息未払前払.利息残高
    
    If (p利息未払前払.利息額増 < 0 Or p利息未払前払.利息額減 < 0) And p利息未払前払.日割日数 > 0 Then '2012/03/07
        G利息未払前払テーブル(j).日割日数 = -p利息未払前払.日割日数     '2012/03/07
    Else                                                                '2012/03/07
        G利息未払前払テーブル(j).日割日数 = p利息未払前払.日割日数      '2012/03/07
    End If                                                              '2012/03/07
    
    G利息未払前払テーブル(j).利率 = p利息未払前払.利率
    G利息未払前払テーブル(j).開始年月日 = p利息未払前払.開始年月日
    G利息未払前払テーブル(j).終了年月日 = p利息未払前払.終了年月日
    G利息未払前払テーブル(j).据置x回目 = p利息未払前払.据置x回目
    
    G利息未払前払テーブル(j).利息期間対象日数 = p利息未払前払.利息期間対象日数  '2014/08/26
    G利息未払前払テーブル(j).利息期間対象額 = p利息未払前払.利息期間対象額      '2014/08/26
    G利息未払前払テーブル(j).利息調整F = p利息未払前払.利息調整F      '2014/08/26
    
    
    '*****以下　中間利払最終日控除の調整処理 を　実行しない場合
    'Exit Sub                        '2016/09/27
    
    '***　中間利払最終日控除の調整処理              2016/09*26
    If p借入計画.利息控除区分 = 4 Then             '2016/09/28
        If p借入計画.利息区分 = XMXA020_区分("利息区分", "利息先払") Then       '2016/89/26
            '**利息先払い
            If G利息未払前払テーブル(j).利息額増 <> 0 Then                      '2016/09/26
                '*前払利息増　発生の時の処理
                If Format(G利息未払前払テーブル(j).返済年月日, "yyyy/mm/dd") _
                    = Format(p借入計画.実行日, "yyyy/mm/dd") Then               '2016/09/26
                    '*実行日の時
                    
                    OLD開始日 = G利息未払前払テーブル(j).開始年月日             '2016/09/26
                    OLD終了日 = G利息未払前払テーブル(j).終了年月日             '2016/09/26
                    NEW開始日 = G利息未払前払テーブル(j).開始年月日             '2016/09/26
                    NEW終了日 = G利息未払前払テーブル(j).終了年月日             '2016/09/26
                Else                                                            '2016/09/26
                    '実行日以外の時
                    OLD開始日 = G利息未払前払テーブル(j).開始年月日             '2016/09/26
                    OLD終了日 = G利息未払前払テーブル(j).終了年月日             '2016/09/26
                    NEW開始日 = DateAdd("d", -1, G利息未払前払テーブル(j).開始年月日)           '2016/09/26
                    NEW終了日 = DateAdd("d", -1, G利息未払前払テーブル(j).終了年月日)           '2016/09/26
                End If                                                          '2016/09/26
                
                G利息未払前払テーブル(j).開始年月日 = NEW開始日                 '2016/09/26
                G利息未払前払テーブル(j).終了年月日 = NEW終了日                 '2016/09/26
                
                WT合計利息額 = G利息未払前払テーブル(j).利息額増                '2016/09/26
                WT合計日数 = G利息未払前払テーブル(j).日割日数                  '2016/09/26
                WT集計利息額 = 0                                                '2016/09/26
            Else                                                                '2016/09/26
                
                '**前払利息減　発生の処理
                'If Format(G利息未払前払テーブル(j).開始年月日, "yyyy/mm/dd") _
                '    = Format(p借入計画.実行日, "yyyy/mm/dd") Then
                'Else
              
                    If Format(OLD開始日, "yyyy/mm/dd") = Format(G利息未払前払テーブル(j).開始年月日, "yyyy/mm/dd") Then
                        G利息未払前払テーブル(j).開始年月日 = NEW開始日             '2016/09/26
                    End If                                                          '2016/09/26
                
                    If Format(OLD終了日, "yyyy/mm/dd") = Format(G利息未払前払テーブル(j).終了年月日, "yyyy/mm/dd") Then
                        G利息未払前払テーブル(j).終了年月日 = NEW終了日             '2016/09/26
                    End If                                                          '2016/09/26
                'End If
              
                G利息未払前払テーブル(j).日割日数 = DateDiff("d", G利息未払前払テーブル(j).開始年月日 _
                                                        , G利息未払前払テーブル(j).終了年月日) + 1      '2016/09/26
                                                        
                If Format(NEW終了日, "yyyy/mm/dd") = Format(G利息未払前払テーブル(j).終了年月日, "yyyy/mm/dd") Then
                    G利息未払前払テーブル(j).利息額減 = WT合計利息額 - WT集計利息額     '2016/09/26
                    G利息未払前払テーブル(j).返済年月日 = NEW終了日             '2016/09/26
                Else                                                            '2016/09/26
                    G利息未払前払テーブル(j).利息額減 = Fix(WT合計利息額 * G利息未払前払テーブル(j).日割日数 / WT合計日数)
                    WT集計利息額 = WT集計利息額 + G利息未払前払テーブル(j).利息額減           '2016/09/26
                End If                                                          '2016/09/26
            End If                                                              '2016/09/26
            
        Else                                                                    '2016/09/26
        '**利息後払い
            If G利息未払前払テーブル(j).利息額減 <> 0 Then                      '2016/09/26
                '未払い利息減　発生の時の処理
                OLD開始日 = G利息未払前払テーブル(j).開始年月日                 '2016/09/26
                OLD終了日 = G利息未払前払テーブル(j).終了年月日                 '2016/09/26
                NEW開始日 = DateAdd("d", -1, G利息未払前払テーブル(j).開始年月日) '2016/09/26
                NEW終了日 = DateAdd("d", -1, G利息未払前払テーブル(j).終了年月日)   '2016/09/26
                
                G利息未払前払テーブル(j).開始年月日 = NEW開始日                 '2016/09/26
                G利息未払前払テーブル(j).終了年月日 = NEW終了日                 '2016/09/26
                
                WT合計利息額 = G利息未払前払テーブル(j).利息額減                '2016/09/26
                WT合計日数 = G利息未払前払テーブル(j).日割日数                  '2016/09/26
                WT集計利息額 = 0                                                '2016/09/26
            Else                                                                '2016/09/26
                '**未払利息増の発生の時の処理
                If Format(OLD開始日, "yyyy/mm/dd") = Format(G利息未払前払テーブル(j).開始年月日, "yyyy/mm/dd") Then
                    G利息未払前払テーブル(j).開始年月日 = NEW開始日             '2016/09/26
                End If                                                          '2016/09/26
                
                If Format(OLD終了日, "yyyy/mm/dd") = Format(G利息未払前払テーブル(j).終了年月日, "yyyy/mm/dd") Then
                    G利息未払前払テーブル(j).終了年月日 = NEW終了日             '2016/09/26
                End If                                                          '2016/09/26
                 
                G利息未払前払テーブル(j).日割日数 = DateDiff("d", G利息未払前払テーブル(j).開始年月日, _
                                                             G利息未払前払テーブル(j).終了年月日) + 1   '2016/09/26
                                                             
                If Format(NEW終了日, "yyyy/mm/dd") = Format(G利息未払前払テーブル(j).終了年月日, "yyyy/mm/dd") Then
                    G利息未払前払テーブル(j).利息額増 = WT合計利息額 - WT集計利息額     '2016/09/26
                    G利息未払前払テーブル(j).返済年月日 = NEW終了日             '2016/09/26
                Else                                                            '2016/09/26
                    G利息未払前払テーブル(j).利息額増 = Fix(WT合計利息額 * G利息未払前払テーブル(j).日割日数 / WT合計日数)
                    WT集計利息額 = WT集計利息額 + G利息未払前払テーブル(j).利息額増         '2016/09/26
                End If                                                          '2016/09/26
            End If                                                              '2016/09/26
            
        End If
        
    End If
    
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_利息未払前払Write_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_利息未払前払Write() でエラー" + vbCrLf + vbCrLf + _
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
' MBD010_利息未払前払明細作成  2011/02/11
'------------------------------------------------
Public Sub MBD010_利息未払前払明細作成()


'
    Dim wRs As ADODB.Recordset
    Dim wstr As String

    Dim j As Integer
    Dim w解約実行日 As Variant      ' 07/02/21 V180
    Dim w返済回数 As Integer        '10/02/27
    Dim w連番 As Integer            '2012/03/06
    
'
    On Error GoTo MBD010_利息未払前払明細作成_ERR
'
    w連番 = 0               '2012/03/06
    
    '**DCDA030_利息未払前払明細 削除
    wstr = ""
    wstr = wstr + "Delete * From DCDA030_利息未払前払明細"
    GDb.Execute wstr



    w返済回数 = 0                   '10/02/27
'
    wstr = ""
    wstr = wstr + "Select * From DCDA030_利息未払前払明細"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        For j = 1 To UBound(G利息未払前払テーブル)
          'If G借入金テーブル(j).元金額 <> 0 Or G借入金テーブル(j).利息額 <> 0 _
          '   Or (G借入金テーブル(j).融資残高 <> 0 And p借入計画マスタ.実行日 = G借入金テーブル(j).実際年月日) _
          '   Or G借入金テーブル(j).保証料 <> 0 Or G借入金テーブル(j).手数料 <> 0 _
          '   Or Format(w解約実行日, "yyyymmdd") = _
          '         Format(G借入金テーブル(j).実際年月日, "yyyymmdd") Then '10/06/16 V195
             
            wRs.AddNew
            
            
                wRs("銀行番号") = G利息未払前払テーブル(j).銀行番号
                wRs("借入番号") = G利息未払前払テーブル(j).借入番号
                wRs("利息区分") = G利息未払前払テーブル(j).利息区分
                wRs("利息計算日数区分") = G利息未払前払テーブル(j).利息計算日数区分
                wRs("返済年月日") = G利息未払前払テーブル(j).返済年月日
                wRs("締年月") = MBA010_締日年月日(G利息未払前払テーブル(j).返済年月日) '11/02/17
                
                If G利息未払前払テーブル(j).据置x回目 = 1 Or G利息未払前払テーブル(j).据置x回目 = 3 Then
                    w連番 = w連番 + 1           '2012/03/06
                    wRs("月毎NO") = w連番           '2012/03/06
                Else                            '2012/03/06
                    wRs("月毎NO") = 0           '2012/03/06
                End If                          '2012/03/06
                
                'wRs("月毎NO") = G利息未払前払テーブル(j).月毎NO
                wRs("元金額") = G利息未払前払テーブル(j).元金額
                wRs("融資残高") = G利息未払前払テーブル(j).融資残高
                
                'If G利息未払前払テーブル(j).据置x回目 = 1 Or G利息未払前払テーブル(j).据置x回目 = 3 Then
                '    wRs("利息計算対象額") = G利息未払前払テーブル(j).元金額
                'Else
                wRs("利息計算対象額") = G利息未払前払テーブル(j).利息計算対象額
                'End If
                
                wRs("利息額増") = G利息未払前払テーブル(j).利息額増
                wRs("利息額減") = G利息未払前払テーブル(j).利息額減
                wRs("利息残高") = G利息未払前払テーブル(j).利息残高
                wRs("日割日数") = G利息未払前払テーブル(j).日割日数
                wRs("利率") = G利息未払前払テーブル(j).利率
                wRs("開始年月日") = G利息未払前払テーブル(j).開始年月日
                wRs("終了年月日") = G利息未払前払テーブル(j).終了年月日
                
                wRs("利息期間対象日数") = G利息未払前払テーブル(j).利息期間対象日数     '2014/08/26
                wRs("利息期間対象額") = G利息未払前払テーブル(j).利息期間対象額         '2014/08/26
                wRs("利息調整F") = G利息未払前払テーブル(j).利息調整F                   '2014/08/26
            wRs.Update
         
        Next
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD010_利息未払前払明細作成_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD010_利息未払前払明細作成() でエラー" + vbCrLf + vbCrLf + _
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
' MRB010_標準平均利率
'------------------------------------------------
Public Sub MRB010_標準平均利率(pTbl As String)
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String, wWhere As String
    
    Dim p借入計画マスタ As MAA910_借入金                                       '5/10/8 V129
    Dim w借入金 As MAA910_借入金
    
    Dim j As Integer, k As Integer
    Dim l As Integer                                                            '11/06/15 V200
    Dim w千円単位 As Integer
    Dim w間隔 As Integer, w回数 As Integer, wcnt As Integer
    Dim w開始cnt As Integer                                                     '11/06/15 V200
    Dim w終了cnt As Integer                                                     '11/06/15 V200
    Dim w日数 As Integer                                                        '11/06/15 V200
      
       
    Dim w基準年月 As Date                                   '2008/02/06 V182
    Dim w開始年月日 As Date
    Dim w年月(12) As Date
    Dim w対象年月 As Date, w対象年月調整 As Date                               'V120
    
    Dim w対象年月日 As Date
    Dim w次回年月日 As Date
    Dim w最終回 As Integer
    
    
    
    Dim w残高開始年月日 As Date                                                 '11/06/15 V200
    Dim w残高終了年月日 As Date                                                 '11/06/15 V200
    Dim w利息残高開始年月日 As Date                                             '11/06/15 V200
    Dim w利息残高終了年月日 As Date                                             '11/06/15 V200
    Dim w借入残 As Double                                                       '11/06/15 V200
    Dim w平均残高合計 As Double, w平均残高(12) As Double                        '11/06/15 V200
    Dim w利息計算平均残高合計 As Double, w利息計算平均残高(12) As Double        '11/06/15 V200
    Dim w平均残高日数合計 As Double, w平均残高日数(12) As Double                '11/06/15Ｖ200
    Dim w平均利息基礎額合計 As Double, w平均利息基礎額(12) As Double            '11/06/15Ｖ200
    
    '***集計範囲のMAXの年月日
    Dim w判定最終年月日 As Date                                                 '11/06/17 V200
    
    '***集計範囲のMinの年月日
    Dim w判定開始年月日 As Date                                                 '11/06/17 V200
    
    Dim w年 As Integer
    Dim w月 As Integer
    Dim w日 As Integer
    Dim w閏年 As Integer
    
    Dim w利息額残高 As Double
    Dim w日数残 As Integer
    Dim w利息基礎額 As Double
    
    Dim w解約実行日 As Variant
    Dim w前払判定日 As Variant          '2012/03/12
    Dim w判定日割日数 As Integer        '2012/03/12
    
    
    
      
    Dim w分母 As String
    Dim w開始年 As String
    Dim w借入計画番号 As String, w金融リストラ As String
    Dim ws基本 As String                                                       '5/8/30 V129
    Dim wベンチャ As String, wベンチャcode As String
    '5/10/9 V129 支店への貸付　リストラ番号
    Dim w期首借入年度 As String, w支店貸付 As String, w全社借入 As String
'
    On Error GoTo MRB010_標準平均利率_ERR
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
    
    
    For j = 0 To wcnt
        w年 = Format(w年月(j), "yyyy")
        w月 = Format(w年月(j), "mm")
        w日 = Format(w年月(j), "dd")
        w日 = G基本情報.決算締日
        
        w閏年 = w年 Mod 4
        
        If w日 >= 29 Then
            If w月 = 2 Then
                If w閏年 = 0 Then
                    w日 = 29
                Else
                    w日 = 28
                End If
            End If
         End If
         
         If w日 = 31 And (w月 = 4 Or w月 = 6 Or w月 = 9 Or w月 = 11) Then
            w日 = 30
         End If
         
         w年月(j) = Format(CStr(w年) & "/" & CStr(w月) & "/" & CStr(w日))
         
    Next
    
    '**集計範囲の最終回の設定
    If w間隔 = 1 Or w間隔 = 3 Then
        w最終回 = 12
    Else
        w最終回 = 10
    End If
    
         
    
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
    wstr = wstr + " Where 手入力区分 = 0"                   ' 07/02/09 V180
    
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
                w平均残高(j) = 0                    '11/06/15 V200
                 
                w利息計算平均残高(j) = 0            '11/06/15 V200
                w平均残高日数(j) = 0                '11/06/15 V200
                w平均利息基礎額(j) = 0              '11/06/15 V200
            Next
            
            w平均残高合計 = 0
            w利息計算平均残高合計 = 0               '11/06/15 V200
            
            w平均残高日数合計 = 0                   '11/06/15 V200
            w平均利息基礎額合計 = 0                 '11/06/15 V200
            
            '***解約算出
            If w金融リストラ > "" And w金融リストラ = p借入計画マスタ.金融リストラ番号 Then
                w解約実行日 = p借入計画マスタ.金融解約実行日
            Else
                w解約実行日 = p借入計画マスタ.解約実行日
            End If
            
            
            
            '** 借入金テーブル セット **
             Call MBD010_借入金テーブル作成(w金融リストラ, p借入計画マスタ) '11/06/15 V200
            
            '*****　メインルーチン
            
            '***うち入れの時　DATEの最後の年月日を、借入計画マスタの最終返済実行日にSET
            '*** 融資残高=0　の時
                                    
            
            For j = 1 To UBound(G借入金テーブル)            '2012/02/29
                If G借入金テーブル(j).融資残高 = 0 Then     '2012/02/29
                    p借入計画マスタ.最終返済実行日 = G借入金テーブル(j).実際年月日  '2012/02/29
                    Exit For                                '2012/02/29
                End If                                      '2012/02/29
            Next                                            '2012/02/29
                                                            
            For j = 1 To UBound(G借入金テーブル)             '11/06/15 V200
            
                If G借入金テーブル(j).実際年月日 > p借入計画マスタ.最終返済実行日 Then  '2012/02/29
                    Exit For                                '2012/02/29
                End If
                
                If G借入金テーブル(j).元金額 = 0 And _
                    G借入金テーブル(j).日割日数 = 0 Then
                    'G借入金テーブル(j).融資残高 = 0 Then    '2012/07/17
                    GoTo 対象外平均残高
                 End If
                 
                
                w日数残 = 0
                w利息額残高 = 0
                
                
                If p借入計画マスタ.利息区分 = "1" Then      '11/06/15 V200
                    l = j                                   '11/06/15 V200
                    
STA1:
                    '***利息先払の時
                    If G借入金テーブル(j).実際年月日 = p借入計画マスタ.最終返済実行日 Then
                        GoTo 対象外平均残高
                    End If
                    
                    If Not IsNull(w解約実行日) And _
                            G借入金テーブル(l + 1).実際年月日 = w解約実行日 Then
                        If G借入金テーブル(l + 1).元金額 = 0 Then
                            G借入金テーブル(l + 1).元金額 = G借入金テーブル(l + 1).融資残高
                        End If
                    End If
                    
                    
                    
                    If (G借入金テーブル(l + 1).元金額 <> 0 Or _
                        G借入金テーブル(l + 1).日割日数 <> 0) _
                        And (G借入金テーブル(l + 1).据置x回目 <> 3) Then
                        
                        If Not IsNull(w解約実行日) And _
                            G借入金テーブル(l + 1).実際年月日 = w解約実行日 _
                             And G借入金テーブル(j).据置x回目 <> 1 And G借入金テーブル(j).据置x回目 <> 3 Then   '2012/03/12
                            'G借入金テーブル(j).利息額 = G借入金テーブル(j).利息額 + G借入金テーブル(l + 1).利息額          '2012/03/14
                            'G借入金テーブル(j).日割日数 = G借入金テーブル(j).日割日数 + G借入金テーブル(l + 1).日割日数    '2012/03/04
                            p借入計画マスタ.最終返済実行日 = w解約実行日
                        End If
                        
                        'If G借入金テーブル(j).実際年月日 = p借入計画マスタ.実行日 And _
                        '   (p借入計画マスタ.利息控除区分 = 1 Or p借入計画マスタ.利息控除区分 = 3) Then
                        '    G借入金テーブル(j).日割日数 = G借入金テーブル(j).日割日数 + 1
                        'End If
                        
                        'If G借入金テーブル(l + 1).実際年月日 = p借入計画マスタ.最終返済実行日 And _
                        '    p借入計画マスタ.最終返済実行日 <> w解約実行日 And _
                        '   (p借入計画マスタ.利息控除区分 = 2 Or p借入計画マスタ.利息控除区分 = 3) Then
                        '    G借入金テーブル(j).日割日数 = G借入金テーブル(j).日割日数 + 1
                        'End If                          '2012/03/23
                        
                        
                        'If Not IsNull(w解約実行日) And G借入金テーブル(l + 1).実際年月日 = w解約実行日 _
                        ' And (p借入計画マスタ.利息控除区分 = 2 Or p借入計画マスタ.利息控除区分 = 3) Then
                        '    G借入金テーブル(j).日割日数 = G借入金テーブル(j).日割日数 + 1
                        'End If
                        
                        
                        '***借入番号　表示
                        p借入計画マスタ.借入番号 = p借入計画マスタ.借入番号
                        
                    
                        w残高開始年月日 = G借入金テーブル(j).実際年月日
                        w残高終了年月日 = DateAdd("D", -1, G借入金テーブル(l + 1).実際年月日)
                    
                        w利息残高開始年月日 = DateAdd("D", 1, G借入金テーブル(j).利息計算年月日)
                        w利息残高終了年月日 = G借入金テーブル(l + 1).利息計算年月日
                        
                        
                        
                           
                        
                        
                        
                        w借入残 = G借入金テーブル(j).融資残高
                    
                        If G借入金テーブル(j).利息計算年月日 = p借入計画マスタ.実行日 Then
                            w利息残高開始年月日 = G借入金テーブル(j).利息計算年月日
                        Else
                            w利息残高開始年月日 = DateAdd("D", 1, G借入金テーブル(j).利息計算年月日)
                        End If
                        
                        
                        
                        
                        '***中間利払いの時　開始年月日　終了年月日　を　1日減ずる
                            p借入計画マスタ.借入番号 = p借入計画マスタ.借入番号 'テスト用
                            If p借入計画マスタ.利息控除区分 = 4 Then              '2016/09/29
                             If p借入計画マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then               '2016/09/28
                                    '**利息先払いの時
                                If Format(p借入計画マスタ.実行日, "yyyy/mm/dd") = Format(w利息残高開始年月日, "yyyy/mm/dd") Then  '2016/09/27
                                Else                                    '2016/09/27
                                        w利息残高開始年月日 = DateAdd("d", -1, w利息残高開始年月日)       '2016/09/29
                                        w利息残高終了年月日 = DateAdd("d", -1, w利息残高終了年月日)     '2016/09/29
                                End If                                              '2016/09/29
                             Else                                                    '2016/09/29
                                '**利息後払いの時
                                    w利息残高開始年月日 = DateAdd("d", -1, w利息残高開始年月日)     '2016/09/29
                                    w利息残高終了年月日 = DateAdd("d", -1, w利息残高終了年月日)     '2016/09/29
                             End If                                              '2016/09/29
                            End If
                        
                        
                        
                        
                        
                        
                        
                        
                        If G借入金テーブル(j).日割日数 < 0 Then
                            G借入金テーブル(j).日割日数 = -G借入金テーブル(j).日割日数
                        End If
                        
                    
                        GoTo STA3
                    
                    Else
                        l = l + 1
                        GoTo STA1
                    End If
                    
                    
                Else
                
                    '***利息後払の時
                    If G借入金テーブル(j).実際年月日 <> p借入計画マスタ.実行日 Then
                        l = j
                        
STA2:

                        If (G借入金テーブル(l - 1).元金額 <> 0 Or _
                            G借入金テーブル(l - 1).日割日数 <> 0 Or _
                            G借入金テーブル(l - 1).実際年月日 = p借入計画マスタ.実行日) And _
                                G借入金テーブル(l - 1).据置x回目 <> 1 Then
                            
                            'If G借入金テーブル(l - 1).実際年月日 = p借入計画マスタ.実行日 And _
                            '    (p借入計画マスタ.利息控除区分 = 1 Or p借入計画マスタ.利息控除区分 = 3) Then
                            '    G借入金テーブル(j).日割日数 = G借入金テーブル(j).日割日数 + 1
                            'End If
                            
                            'If G借入金テーブル(j).実際年月日 = p借入計画マスタ.最終返済実行日 And _
                            '    (p借入計画マスタ.利息控除区分 = 2 Or p借入計画マスタ.利息控除区分 = 3) Then
                            '    G借入金テーブル(j).日割日数 = G借入金テーブル(j).日割日数 + 1
                            'End If
                            
                            'If Not IsNull(w解約実行日) And G借入金テーブル(j).実際年月日 = w解約実行日 _
                            '    And (p借入計画マスタ.利息控除区分 = 2 Or p借入計画マスタ.利息控除区分 = 3) Then
                            '    G借入金テーブル(j).日割日数 = G借入金テーブル(j).日割日数 + 1
                            'End If
                            
                            
                            w残高開始年月日 = G借入金テーブル(l - 1).実際年月日
                            w残高終了年月日 = DateAdd("D", -1, G借入金テーブル(j).実際年月日)
                        
                            w利息残高開始年月日 = DateAdd("D", 1, G借入金テーブル(l - 1).利息計算年月日)
                            w利息残高終了年月日 = G借入金テーブル(j).利息計算年月日
                            
                            
                            
                            
                                 
                            
                            
                            
                            
                            
                            w借入残 = G借入金テーブル(l - 1).融資残高
                        
                            If G借入金テーブル(l - 1).実際年月日 = p借入計画マスタ.実行日 Then
                                w利息残高開始年月日 = G借入金テーブル(l - 1).利息計算年月日
                            Else
                                w利息残高開始年月日 = DateAdd("D", 1, G借入金テーブル(l - 1).利息計算年月日)
                            End If
                            
                            
                            
                            
                            
                            '***中間利払いの時　開始年月日　終了年月日　を　1日減ずる
                            p借入計画マスタ.借入番号 = p借入計画マスタ.借入番号 'テスト用
                            If p借入計画マスタ.利息控除区分 = 4 Then              '2016/09/29
                             If p借入計画マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then               '2016/09/28
                                    '**利息先払いの時
                                If Format(p借入計画マスタ.実行日, "yyyy/mm/dd") = Format(w利息残高開始年月日, "yyyy/mm/dd") Then  '2016/09/27
                                Else                                    '2016/09/27
                                        w利息残高開始年月日 = DateAdd("d", -1, w利息残高開始年月日)       '2016/09/29
                                        w利息残高終了年月日 = DateAdd("d", -1, w利息残高終了年月日)     '2016/09/29
                                End If                                              '2016/09/29
                             Else                                                    '2016/09/29
                                '**利息後払いの時
                                If Format(p借入計画マスタ.実行日, "yyyy/mm/dd") = Format(w利息残高開始年月日, "yyyy/mm/dd") Then  '2016/09/27
                                    w利息残高終了年月日 = DateAdd("d", -1, w利息残高終了年月日)     '2016/09/29
                                Else
                                    w利息残高開始年月日 = DateAdd("d", -1, w利息残高開始年月日)     '2016/09/29
                                    w利息残高終了年月日 = DateAdd("d", -1, w利息残高終了年月日)     '2016/09/29
                                End If
                             End If                                              '2016/09/29
                            End If
                       
                            
                            
                            
                            
                            
                            
                        
                            GoTo STA3
                        
                        Else
                            l = l - 1
                            GoTo STA2
                        End If
                        
                    Else
                        GoTo 対象外平均残高
                    End If
                End If
                
STA3:

    
                '***平均残高計算
                
                 '***平均残高計算
                w対象年月日 = w残高開始年月日
                
STA5:
                w次回年月日 = MRB010_次回年月日算出(w対象年月日)
                
                If w残高終了年月日 > w次回年月日 Then
                    
                    w日数 = DateDiff("D", w対象年月日, w次回年月日) + 1
                    
                    
                Else
                    '***最終回
                    
                    w日数 = DateDiff("D", w対象年月日, w残高終了年月日) + 1
                    
                    
                End If
                
                '*平均残高集計
                'For k = 1 To w最終回
                '    If w次回年月日 > w年月(k - 1) And w次回年月日 <= w年月(k) Then
                '        'w平均残高(k) = w平均残高(k) + Round(w借入残 * w日数)
                '        Exit For
                '    End If
                'Next
                
                If w残高終了年月日 > w次回年月日 Then
                    w対象年月日 = DateAdd("D", 1, w次回年月日)
                    GoTo STA5
                End If
                      
                  
'対象外平均残高:


                '***利息平均残高計算
                
                w対象年月日 = w利息残高開始年月日
                
STA6:
                w次回年月日 = MRB010_次回年月日算出(w対象年月日)
                
                If w利息残高終了年月日 > w次回年月日 Then
                    
                  If w対象年月日 = p借入計画マスタ.実行日 And _
                     (p借入計画マスタ.利息控除区分 = 1 Or p借入計画マスタ.利息控除区分 = 3) Then    '2012/07/06
                    w日数 = DateDiff("D", w対象年月日, w次回年月日)     '2012/07/06
                  Else                                                  '2012/07/06
                    
                    w日数 = DateDiff("D", w対象年月日, w次回年月日) + 1
                  End If                                                '2012/07/06
                  
                    
                    
                Else
                    '***最終回
                    
                    w日数 = DateDiff("D", w対象年月日, w利息残高終了年月日) + 1
                    
                    
                    
                End If
                
                '*利息計算平均残高集計
                w日数残 = w日数残 + w日数
                
                '*利息前払、解約の時調整
                w判定日割日数 = G借入金テーブル(j).日割日数         '2012/03/12
                If p借入計画マスタ.利息区分 = "1" And Not IsNull(w解約実行日) Then  '2012/03/12
                    w前払判定日 = DateAdd("d", G借入金テーブル(j).日割日数, G借入金テーブル(j).利息計算年月日)  '2012/03/12
                    If w前払判定日 > w解約実行日 Then               '2012/03/12
                        w判定日割日数 = DateDiff("d", G借入金テーブル(j).利息計算年月日, w解約実行日) '2012/03/12
                    End If                                          '2012/03/12
                End If                                              '2012/03/12
                
                    
                'If w日数残 >= G借入金テーブル(j).日割日数 Then
                If w日数残 >= w判定日割日数 Then                    '2012/03/12
                    w利息基礎額 = G借入金テーブル(j).利息額 - w利息額残高
                    
                    '*利息先払の解約の時の調整
                    If p借入計画マスタ.利息区分 = "1" Then                   '2012/03/14
                      If Not IsNull(w解約実行日) Then                      '2012/03/14
                        l = j                                               '2012/03/14
                        
                        For l = j To UBound(G借入金テーブル)                '2012/03/14
                            If G借入金テーブル(l + 1).利息額 <> 0 Or _
                               G借入金テーブル(l + 1).実際年月日 = w解約実行日 Then     '2012/07/03
                                If G借入金テーブル(l + 1).実際年月日 = w解約実行日 Then '2012/03/14
                            
                                    w利息基礎額 = w利息基礎額 + G借入金テーブル(l + 1).利息額   '2012/03/14
                                    Exit For                                '2012/03/14
                                Else                                        '2012/03/14
                                    Exit For                                '2012/03/14
                                End If                                                      '2012/03/14
                            End If                                          '2012/03/14
                        Next                                                '2014/03/14
                        
                      End If                                                '2012/03/14
                    End If                                                          '2012/03/14
                    
                Else
                    w利息基礎額 = Fix(G借入金テーブル(j).利息額 * w日数 / G借入金テーブル(j).日割日数)  '2012/03/07
                End If
                    
                w利息額残高 = w利息額残高 + w利息基礎額
                
                For k = 1 To w最終回
                    If w次回年月日 > w年月(k - 1) And w次回年月日 <= w年月(k) Then
                        
                        'w利息計算平均残高(k) = w利息計算平均残高(k) + Round(w借入残 * w日数)
                    
                        w平均利息基礎額(k) = w平均利息基礎額(k) + w利息基礎額
                        Exit For
                    End If
                Next
                
                If w利息残高終了年月日 > w次回年月日 Then
                    w対象年月日 = DateAdd("D", 1, w次回年月日)
                    GoTo STA6
                End If
  
                
                
対象外平均残高:

対象外利息平均残高:

            
            
            Next
            
            
            
            
            '***標準　平均残高　利息計算平均残高　算出
            '** 借入金テーブル セット **
             Call MBD010_借入金テーブル作成(w金融リストラ, p借入計画マスタ) '11/06/15 V200
            
            '*****　メインルーチン
            
            For j = 1 To UBound(G借入金テーブル)            '2012/02/29
                If G借入金テーブル(j).融資残高 = 0 Then     '2012/02/29
                    p借入計画マスタ.最終返済実行日 = G借入金テーブル(j).実際年月日  '2012/02/29
                    Exit For                                '2012/02/29
                End If                                      '2012/02/29
            Next                                            '2012/02/29
            
            
            For j = 1 To UBound(G借入金テーブル)                '11/06/15 V200
            
                If G借入金テーブル(j).実際年月日 > p借入計画マスタ.最終返済実行日 Then  '2012/02/29
                    Exit For                                    '2012/02/29
                End If                                          '2012/02/29
                
                If Not IsNull(w解約実行日) And G借入金テーブル(j).実際年月日 = w解約実行日 _
                    And (p借入計画マスタ.利息控除区分 = 2 Or p借入計画マスタ.利息控除区分 = 3) Then
                    GoTo Ok1                                    '2012/07/17
                End If                                          '2012/07/17
                

                If G借入金テーブル(j).元金額 = 0 And _
                    G借入金テーブル(j).日割日数 = 0 Then
                    GoTo 対象外平均残高HH
                End If
                
Ok1:
                 
                
                w日数残 = 0
                w利息額残高 = 0
                
                
                If p借入計画マスタ.利息区分 = "1" Then      '11/06/15 V200
                    l = j                                   '11/06/15 V200
                    
STA1HH:
                    '***利息先払の時
                    If G借入金テーブル(j).実際年月日 = p借入計画マスタ.最終返済実行日 Then
                        GoTo 対象外平均残高HH
                    End If
                    
                    If Not IsNull(w解約実行日) And _
                            G借入金テーブル(l + 1).実際年月日 = w解約実行日 Then
                        If G借入金テーブル(l + 1).元金額 = 0 Then
                            G借入金テーブル(l + 1).元金額 = G借入金テーブル(l + 1).融資残高
                        End If
                    End If
                    
                    
                    
                    If (G借入金テーブル(l + 1).元金額 <> 0 Or _
                        G借入金テーブル(l + 1).日割日数 <> 0) Then
                        'And (G借入金テーブル(l + 1).据置X回目 <> 3) Then
                        
                        If Not IsNull(w解約実行日) And _
                            G借入金テーブル(l + 1).実際年月日 = w解約実行日 Then
                            G借入金テーブル(j).利息額 = G借入金テーブル(j).利息額 + G借入金テーブル(l + 1).利息額
                            G借入金テーブル(j).日割日数 = G借入金テーブル(j).日割日数 + G借入金テーブル(l + 1).日割日数
                            p借入計画マスタ.最終返済実行日 = w解約実行日
                        End If
                        
                        If G借入金テーブル(j).実際年月日 = p借入計画マスタ.実行日 And _
                           (p借入計画マスタ.利息控除区分 = 1 Or p借入計画マスタ.利息控除区分 = 3) Then
                            G借入金テーブル(j).日割日数 = G借入金テーブル(j).日割日数 + 1
                        End If
                        
                        If G借入金テーブル(l + 1).実際年月日 = p借入計画マスタ.最終返済実行日 And _
                           (p借入計画マスタ.利息控除区分 = 2 Or p借入計画マスタ.利息控除区分 = 3) Then
                            G借入金テーブル(j).日割日数 = G借入金テーブル(j).日割日数 + 1
                        End If
                        
                        
                        'If Not IsNull(w解約実行日) And G借入金テーブル(l + 1).実際年月日 = w解約実行日 _
                        ' And (p借入計画マスタ.利息控除区分 = 2 Or p借入計画マスタ.利息控除区分 = 3) Then
                        '    G借入金テーブル(j).日割日数 = G借入金テーブル(j).日割日数 + 1
                        'End If
                        
                    
                        w残高開始年月日 = G借入金テーブル(j).実際年月日
                        w残高終了年月日 = DateAdd("D", -1, G借入金テーブル(l + 1).実際年月日)
                    
                        w利息残高開始年月日 = DateAdd("D", 1, G借入金テーブル(j).利息計算年月日)
                        w利息残高終了年月日 = G借入金テーブル(l + 1).利息計算年月日
                        w借入残 = G借入金テーブル(j).融資残高
                    
                        If G借入金テーブル(j).利息計算年月日 = p借入計画マスタ.実行日 Then
                            w利息残高開始年月日 = G借入金テーブル(j).利息計算年月日
                        Else
                            w利息残高開始年月日 = DateAdd("D", 1, G借入金テーブル(j).利息計算年月日)
                        End If
                        
                        
                        
                        
                        
                        '***中間利払いの時　開始年月日　終了年月日　を　1日減ずる
                            p借入計画マスタ.借入番号 = p借入計画マスタ.借入番号 'テスト用
                            If p借入計画マスタ.利息控除区分 = 4 Then              '2016/09/29
                             If p借入計画マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then               '2016/09/28
                                '**利息先払いの時
                                If Format(p借入計画マスタ.実行日, "yyyy/mm/dd") = Format(w利息残高開始年月日, "yyyy/mm/dd") Then  '2016/09/27
                                Else                                    '2016/09/27
                                        w利息残高開始年月日 = DateAdd("d", -1, w利息残高開始年月日)       '2016/09/29
                                        w利息残高終了年月日 = DateAdd("d", -1, w利息残高終了年月日)     '2016/09/29
                                End If                                              '2016/09/29
                             Else                                                    '2016/09/29
                                '**利息後払いの時
                                If Format(p借入計画マスタ.実行日, "yyyy/mm/dd") = Format(w利息残高開始年月日, "yyyy/mm/dd") Then  '2016/09/27
                                    w利息残高終了年月日 = DateAdd("d", -1, w利息残高終了年月日)     '2016/09/29
                                Else
                                    w利息残高開始年月日 = DateAdd("d", -1, w利息残高開始年月日)     '2016/09/29
                                    w利息残高終了年月日 = DateAdd("d", -1, w利息残高終了年月日)     '2016/09/29
                                End If
                             End If                                              '2016/09/29
                            End If
                       
                        
                        
                        
                        
                        If G借入金テーブル(j).日割日数 < 0 Then
                            G借入金テーブル(j).日割日数 = -G借入金テーブル(j).日割日数
                        End If
                        
                    
                        GoTo STA3HH
                    
                    Else
                        l = l + 1
                        GoTo STA1HH
                    End If
                    
                    
                Else
                
                    '***利息後払の時
                    If G借入金テーブル(j).実際年月日 <> p借入計画マスタ.実行日 Then
                        l = j
                        
STA2HH:

                        If (G借入金テーブル(l - 1).元金額 <> 0 Or _
                            G借入金テーブル(l - 1).日割日数 <> 0 Or _
                            G借入金テーブル(l - 1).実際年月日 = p借入計画マスタ.実行日) Then
                                'G借入金テーブル(l - 1).据置X回目 <> 1 Then
                            
                            If G借入金テーブル(l - 1).実際年月日 = p借入計画マスタ.実行日 And _
                                (p借入計画マスタ.利息控除区分 = 1 Or p借入計画マスタ.利息控除区分 = 3) Then
                                G借入金テーブル(j).日割日数 = G借入金テーブル(j).日割日数 + 1
                            End If
                            
                            If G借入金テーブル(j).実際年月日 = p借入計画マスタ.最終返済実行日 And _
                                (p借入計画マスタ.利息控除区分 = 2 Or p借入計画マスタ.利息控除区分 = 3) Then
                                G借入金テーブル(j).日割日数 = G借入金テーブル(j).日割日数 + 1
                            End If
                            
                            If Not IsNull(w解約実行日) And G借入金テーブル(j).実際年月日 = w解約実行日 _
                                And (p借入計画マスタ.利息控除区分 = 2 Or p借入計画マスタ.利息控除区分 = 3) Then
                                G借入金テーブル(j).日割日数 = G借入金テーブル(j).日割日数 + 1
                            End If
                            
                            
                            w残高開始年月日 = G借入金テーブル(l - 1).実際年月日
                            w残高終了年月日 = DateAdd("D", -1, G借入金テーブル(j).実際年月日)
                        
                            w利息残高開始年月日 = DateAdd("D", 1, G借入金テーブル(l - 1).利息計算年月日)
                            w利息残高終了年月日 = G借入金テーブル(j).利息計算年月日
                            w借入残 = G借入金テーブル(l - 1).融資残高
                        
                            If G借入金テーブル(l - 1).実際年月日 = p借入計画マスタ.実行日 Then
                                w利息残高開始年月日 = G借入金テーブル(l - 1).利息計算年月日
                            Else
                                w利息残高開始年月日 = DateAdd("D", 1, G借入金テーブル(l - 1).利息計算年月日)
                            End If
                            
                            
                            
                            
                            
                            
                            '***中間利払いの時　開始年月日　終了年月日　を　1日減ずる
                            p借入計画マスタ.借入番号 = p借入計画マスタ.借入番号     'テスト用
                            If p借入計画マスタ.利息控除区分 = 4 Then              '2016/09/29
                             If p借入計画マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then               '2016/09/28
                                    '**利息先払いの時
                                If Format(p借入計画マスタ.実行日, "yyyy/mm/dd") = Format(w利息残高開始年月日, "yyyy/mm/dd") Then  '2016/09/27
                                
                                Else                                    '2016/09/27
                                        w利息残高開始年月日 = DateAdd("d", -1, w利息残高開始年月日)       '2016/09/29
                                        w利息残高終了年月日 = DateAdd("d", -1, w利息残高終了年月日)     '2016/09/29
                                End If                                              '2016/09/29
                             Else                                                    '2016/09/29
                                '**利息後払いの時
                                If Format(p借入計画マスタ.実行日, "yyyy/mm/dd") = Format(w利息残高開始年月日, "yyyy/mm/dd") Then  '2016/09/27
                                    w利息残高終了年月日 = DateAdd("d", -1, w利息残高終了年月日)     '2016/09/29
                                Else
                                    w利息残高開始年月日 = DateAdd("d", -1, w利息残高開始年月日)     '2016/09/29
                                    w利息残高終了年月日 = DateAdd("d", -1, w利息残高終了年月日)     '2016/09/29
                                End If
                             End If                                              '2016/09/29
                            End If
                       
                            
                            
                            
                            
                        
                            GoTo STA3HH
                        
                        Else
                            l = l - 1
                            GoTo STA2HH
                        End If
                        
                    Else
                        GoTo 対象外平均残高HH
                    End If
                End If
                
STA3HH:

    
                '***平均残高計算
                
                 '***平均残高計算
                w対象年月日 = w残高開始年月日
                
STA5HH:
                w次回年月日 = MRB010_次回年月日算出(w対象年月日)
                
                If w残高終了年月日 > w次回年月日 Then
                    
                    w日数 = DateDiff("D", w対象年月日, w次回年月日) + 1
                    
                    
                Else
                    '***最終回
                    
                    w日数 = DateDiff("D", w対象年月日, w残高終了年月日) + 1
                    
                    
                End If
                
                '*平均残高集計
                For k = 1 To w最終回
                    If w次回年月日 > w年月(k - 1) And w次回年月日 <= w年月(k) Then
                        w平均残高(k) = w平均残高(k) + Fix(w借入残 * w日数 + 0.5) '2012/02/29
                        Exit For
                    End If
                Next
                
                If w残高終了年月日 > w次回年月日 Then
                    w対象年月日 = DateAdd("D", 1, w次回年月日)
                    GoTo STA5HH
                End If
                      
                  
'対象外平均残高HH:


                '***利息平均残高計算
                
                w対象年月日 = w利息残高開始年月日
                
STA6HH:
                w次回年月日 = MRB010_次回年月日算出(w対象年月日)
                
                If w利息残高終了年月日 > w次回年月日 Then
                    
                  
                    w日数 = DateDiff("D", w対象年月日, w次回年月日) + 1
                    
                    
                Else
                    '***最終回
                    
                    w日数 = DateDiff("D", w対象年月日, w利息残高終了年月日) + 1
                    
                    
                    
                End If
                
                '*利息計算平均残高集計
                w日数残 = w日数残 + w日数
                    
                If w日数残 >= G借入金テーブル(j).日割日数 Then
                    w利息基礎額 = G借入金テーブル(j).利息額 - w利息額残高
                Else
                    w利息基礎額 = Fix(G借入金テーブル(j).利息額 * w日数 / G借入金テーブル(j).日割日数)  '2012/03/07
                End If
                    
                w利息額残高 = w利息額残高 + w利息基礎額
                
                For k = 1 To w最終回
                    If w次回年月日 > w年月(k - 1) And w次回年月日 <= w年月(k) Then
                        
                        w利息計算平均残高(k) = w利息計算平均残高(k) + Fix(w借入残 * w日数 + 0.5) '2012/02/29
                    
                        'w平均利息基礎額(k) = w平均利息基礎額(k) + w利息基礎額
                        Exit For
                    End If
                Next
                
                If w利息残高終了年月日 > w次回年月日 Then
                    w対象年月日 = DateAdd("D", 1, w次回年月日)
                    GoTo STA6HH
                End If
  
                
                
対象外平均残高HH:

対象外利息平均残高HH:

            
            
            Next
            
            
            
            
            
            
            
              
             For k = 1 To wcnt                  '11/06/15 V200
                w平均残高日数(k) = DateDiff("D", w年月(k - 1), w年月(k)) '11/06/15 V200
             Next                               '11/06/15 V200
             
             
             For k = 1 To wcnt
                w平均残高合計 = w平均残高合計 + w平均残高(k)
                
                w利息計算平均残高合計 = w利息計算平均残高合計 + w利息計算平均残高(k)    '11/06/15 V200
                w平均残高日数合計 = w平均残高日数合計 + w平均残高日数(k)    '11/06/15 V200
                w平均利息基礎額合計 = w平均利息基礎額合計 + w平均利息基礎額(k)  '11/06/15 V200
             Next
              
             '***** 平均残高等による　DCDA010_借入残高推移表結果２　の　更新処理
             If w平均残高合計 <> 0 Or _
                w利息計算平均残高合計 <> 0 Or _
                w平均利息基礎額合計 <> 0 Then
                
                wstr2 = ""
                wstr2 = wstr2 + "Select * From DCDA010_借入残高推移表結果２"
                wstr2 = wstr2 + " Where 借入番号='" & p借入計画マスタ.借入番号 & "'"
                Call AdoRecordsetOpen(GDb, wRs2, wstr2)
                If wRs2.eof Then
                    wRs2.AddNew
                    wRs2("借入番号") = p借入計画マスタ.借入番号
                End If
                
                    wRs2("平均残高合計") = w平均残高合計
                    
                    wRs2("利息計算平均残高合計") = w利息計算平均残高合計    '11/06/15 V200
                    wRs2("平均残高日数合計") = w平均残高日数合計            '11/06/15 V200
                    wRs2("平均利息基礎額合計") = w平均利息基礎額合計        '11/06/15 V200
                    
                    For k = 1 To wcnt
                        wRs2("平均残高_" + CStr(Format(k, "00"))) = w平均残高(k)    '11/06/15 V200
                        
                        wRs2("利息計算平均残高_" + CStr(Format(k, "00"))) = w利息計算平均残高(k)
                        wRs2("平均残高日数_" + CStr(Format(k, "00"))) = w平均残高日数(k)
                        wRs2("平均利息基礎額_" + CStr(Format(k, "00"))) = w平均利息基礎額(k)
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
MRB010_標準平均利率_ERR:
    pERR_MES = pPROGRAM_ID + "/ MRB010_標準平均利率() でエラー" + vbCrLf + vbCrLf + _
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
' MRB010_手入力平均利率
'------------------------------------------------
Public Sub MRB010_手入力平均利率(pTbl As String)
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String, wWhere As String
    
    Dim p借入計画マスタ As MAA910_借入金                                       '5/10/8 V129
    Dim w借入金 As MAA910_借入金
    
    Dim j As Integer, k As Integer
    Dim l As Integer                                                            '11/06/15 V200
    Dim w千円単位 As Integer
    Dim w間隔 As Integer, w回数 As Integer, wcnt As Integer
    Dim w開始cnt As Integer                                                     '11/06/15 V200
    Dim w終了cnt As Integer                                                     '11/06/15 V200
    Dim w日数 As Integer                                                        '11/06/15 V200
    Dim p日数 As Integer                                    '2012/07/13
       
    Dim w基準年月 As Date                                   '2008/02/06 V182
    Dim w開始年月日 As Date
    Dim w年月(12) As Date
    Dim w対象年月 As Date, w対象年月調整 As Date                               'V120
    
    
    Dim w対象年月日 As Date
    Dim w次回年月日 As Date
    Dim w最終回 As Integer
    
    
    
    Dim w残高開始年月日 As Date                                                 '11/06/15 V200
    Dim w残高終了年月日 As Date                                                 '11/06/15 V200
    Dim w利息残高開始年月日 As Date                                             '11/06/15 V200
    Dim w利息残高終了年月日 As Date                                             '11/06/15 V200
    
    '***利息額按分用                                                            ’2012/08/21
    Dim wx利息残高開始年月日 As Date                                            '2012/08/21
    Dim wx利息残高終了年月日 As Date                                            '2012/08/21
    Dim wx日数 As Integer                                                       '2012/08/21
    Dim wx日割日数 As Integer                                                   '2012/08/21
    Dim wx対象年月日 As Date                                                    '2012/08/21
    Dim wx次回年月日 As Date                                                    '2012/08/21
    Dim wx日数残 As Integer                                                     '2012/08/21
    
    
    
    
    Dim w借入残 As Double                                                       '11/06/15 V200
    Dim w平均残高合計 As Double, w平均残高(12) As Double                        '11/06/15 V200
    Dim w利息計算平均残高合計 As Double, w利息計算平均残高(12) As Double        '11/06/15 V200
    Dim w平均残高日数合計 As Double, w平均残高日数(12) As Double                '11/06/15Ｖ200
    Dim w平均利息基礎額合計 As Double, w平均利息基礎額(12) As Double            '11/06/15Ｖ200
    
    '***集計範囲のMAXの年月日
    Dim w判定最終年月日 As Date                                                 '11/06/17 V200
    
    '***集計範囲のMinの年月日
    Dim w判定開始年月日 As Date                                                 '11/06/17 V200
    
    Dim w年 As Integer
    Dim w月 As Integer
    Dim w日 As Integer
    Dim w閏年 As Integer
    
    Dim w利息額残高 As Double
    Dim w日数残 As Integer
    Dim w利息基礎額 As Double
    
    
      
    Dim w分母 As String
    Dim w開始年 As String
    Dim w借入計画番号 As String, w金融リストラ As String
    Dim ws基本 As String                                                       '5/8/30 V129
    Dim wベンチャ As String, wベンチャcode As String
    '5/10/9 V129 支店への貸付　リストラ番号
    Dim w期首借入年度 As String, w支店貸付 As String, w全社借入 As String
'
    On Error GoTo MRB010_手入力平均利率_ERR
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
    
    
    For j = 0 To wcnt
        w年 = Format(w年月(j), "yyyy")
        w月 = Format(w年月(j), "mm")
        w日 = Format(w年月(j), "dd")
        w日 = G基本情報.決算締日
        
        w閏年 = w年 Mod 4
        
        If w日 >= 29 Then
            If w月 = 2 Then
                If w閏年 = 0 Then
                    w日 = 29
                Else
                    w日 = 28
                End If
            End If
         End If
         
         If w日 = 31 And (w月 = 4 Or w月 = 6 Or w月 = 9 Or w月 = 11) Then
            w日 = 30
         End If
         
         w年月(j) = Format(CStr(w年) & "/" & CStr(w月) & "/" & CStr(w日))
         
    Next
    
    '**集計範囲の最終回の設定
    If w間隔 = 1 Or w間隔 = 3 Then
        w最終回 = 12
    Else
        w最終回 = 10
    End If
    
    
         
    
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
                w平均残高(j) = 0                    '11/06/15 V200
                 
                w利息計算平均残高(j) = 0            '11/06/15 V200
                w平均残高日数(j) = 0                '11/06/15 V200
                w平均利息基礎額(j) = 0              '11/06/15 V200
            Next
            
            w平均残高合計 = 0
            w利息計算平均残高合計 = 0               '11/06/15 V200
            
            w平均残高日数合計 = 0                   '11/06/15 V200
            w平均利息基礎額合計 = 0                 '11/06/15 V200
            
            
            '** 借入金テーブル セット **
             Call MBD010_借入金入力明細Read(p借入計画マスタ)  '11/06/15 V200
             
             p借入計画マスタ.借入番号 = p借入計画マスタ.借入番号
            
            '*****　メインルーチン
            
            For j = 1 To UBound(G借入金入力)                 '11/06/15 V200
                If G借入金入力(j).日割日数 = 0 And G借入金入力(j).利息額 = 0 _
                   And G借入金入力(j).元金 = 0 Then         '2012/07/17
                    GoTo 対象外平均残高
                End If
                
                
                
                
                w日数残 = 0
                wx日数残 = 0                               '2012/08/21
                
                w利息額残高 = 0
                
                
                
                If p借入計画マスタ.利息区分 = "1" Then      '11/06/15 V200
                    l = j                                   '11/06/15 V200
                    
STA1:
                    '***利息先払の時
                    If G借入金入力(j).借入返済年月日 = p借入計画マスタ.最終返済実行日 Then
                        GoTo 対象外平均残高
                    End If
                    
                    
                    If G借入金入力(l + 1).元金 <> 0 Or _
                       G借入金入力(l + 1).日割日数 <> 0 Or G借入金入力(l + 1).利息額 <> 0 Then
                        
                        'If G借入金入力(j).借入返済年月日 = p借入計画マスタ.実行日 And _
                        '   G借入金入力(j).日割日数 <> 0 And _
                        '   (p借入計画マスタ.利息控除区分 = 1 Or p借入計画マスタ.利息控除区分 = 3) Then
                        '    G借入金入力(j).日割日数 = G借入金入力(j).日割日数 + 1
                        'End If
                        
                        'If G借入金入力(l + 1).借入返済年月日 = p借入計画マスタ.最終返済実行日 And _
                        '   G借入金入力(j).日割日数 <> 0 And _
                        '   (p借入計画マスタ.利息控除区分 = 2 Or p借入計画マスタ.利息控除区分 = 3) Then
                        '    G借入金入力(j).日割日数 = G借入金入力(j).日割日数 + 1
                        'End If
                        
                        
                    
                        w残高開始年月日 = G借入金入力(j).借入返済年月日
                        w残高終了年月日 = DateAdd("D", -1, G借入金入力(l + 1).借入返済年月日)
                    
                        w利息残高開始年月日 = DateAdd("D", 1, G借入金入力(j).利息計算年月日)
                        w利息残高終了年月日 = G借入金入力(l + 1).利息計算年月日
                        
                        
                        
                        
                            
                        'wx利息残高開始年月日 = w利息残高開始年月日              '2012/08/21
                        'wx利息残高終了年月日 = DateAdd("D", G借入金入力(j).日割日数 - 1, wx利息残高開始年月日) '2012/08/21
                        
                        
                        w借入残 = G借入金入力(j).融資残高
                    
                        If G借入金入力(j).借入返済年月日 = p借入計画マスタ.実行日 Then  '2012/08/28
                            w利息残高開始年月日 = G借入金入力(j).利息計算年月日
                        Else
                            w利息残高開始年月日 = DateAdd("D", 1, G借入金入力(j).利息計算年月日)
                        End If
                        
                        
                        
                        
                        
                        '***中間利払いの時　開始年月日　終了年月日　を　1日減ずる
                            p借入計画マスタ.借入番号 = p借入計画マスタ.借入番号 'テスト用
                            If p借入計画マスタ.利息控除区分 = 4 Then              '2016/09/29
                             If p借入計画マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then               '2016/09/28
                                    '**利息先払いの時
                                If Format(p借入計画マスタ.実行日, "yyyy/mm/dd") = Format(w利息残高開始年月日, "yyyy/mm/dd") Then  '2016/09/27
                                Else                                    '2016/09/27
                                        w利息残高開始年月日 = DateAdd("d", -1, w利息残高開始年月日)       '2016/09/29
                                        w利息残高終了年月日 = DateAdd("d", -1, w利息残高終了年月日)     '2016/09/29
                                End If                                              '2016/09/29
                             Else                                                    '2016/09/29
                                '**利息後払いの時
                                If Format(p借入計画マスタ.実行日, "yyyy/mm/dd") = Format(w利息残高開始年月日, "yyyy/mm/dd") Then  '2016/09/27
                                    w利息残高終了年月日 = DateAdd("d", -1, w利息残高終了年月日)     '2016/09/29
                                Else
                                    w利息残高開始年月日 = DateAdd("d", -1, w利息残高開始年月日)     '2016/09/29
                                    w利息残高終了年月日 = DateAdd("d", -1, w利息残高終了年月日)     '2016/09/29
                                End If
                             End If                                              '2016/09/29
                            End If
                       
                        
                        
                        
                        
                        
                        
                        
                        
                        wx利息残高開始年月日 = w利息残高開始年月日              '2012/08/21
                        
                        If G借入金入力(j).借入返済年月日 = p借入計画マスタ.実行日 And _
                            (p借入計画マスタ.利息控除区分 = 1 Or p借入計画マスタ.利息控除区分 = 3) Then '2012/08/28
                            wx利息残高開始年月日 = DateAdd("D", 1, wx利息残高開始年月日)    '2012/08/28
                        End If                                                              '2012/08/28
                        
                        wx利息残高終了年月日 = DateAdd("D", G借入金入力(j).日割日数 - 1, wx利息残高開始年月日) '2012/08/21
                        
                        
                        wx日割日数 = G借入金入力(j).日割日数                    '2012/08/21
                        
                        
                        If G借入金入力(j).日割日数 = 0 Then
                            G借入金入力(j).日割日数 = DateDiff("D", w利息残高開始年月日, w利息残高終了年月日) + 1
                            wx日割日数 = DateDiff("D", wx利息残高開始年月日, wx利息残高終了年月日) + 1 '2012/08/21
                            
                        End If
                        
                    
                        GoTo STA3
                    
                    Else
                        l = l + 1
                        GoTo STA1
                    End If
                    
                    
                Else
                
                    '***利息後払の時
                    If G借入金入力(j).借入返済年月日 <> p借入計画マスタ.実行日 Then
                        l = j
                        
STA2:
                        If j = 1 And G借入金入力(j).借入返済年月日 <> p借入計画マスタ.実行日 Then
                            G借入金入力(0).借入返済年月日 = p借入計画マスタ.実行日
                            G借入金入力(0).利息計算年月日 = p借入計画マスタ.実行日
                            G借入金入力(0).元金 = 0
                            G借入金入力(0).日割日数 = 0
                        End If
                        
                        
                        'If G借入金入力(l - 1).元金 <> 0 Or
                        If G借入金入力(l - 1).日割日数 <> 0 Or _
                            G借入金入力(l - 1).利息額 <> 0 Or _
                            G借入金入力(l - 1).借入返済年月日 = p借入計画マスタ.実行日 Then
                            
                            'If G借入金入力(l - 1).借入返済年月日 = p借入計画マスタ.実行日 And _
                            '   G借入金入力(l - 1).日割日数 <> 0 And _
                            '    (p借入計画マスタ.利息控除区分 = 1 Or p借入計画マスタ.利息控除区分 = 3) Then
                            '    G借入金入力(j).日割日数 = G借入金入力(j).日割日数 + 1
                            'End If
                            
                            'If G借入金入力(j).借入返済年月日 = p借入計画マスタ.最終返済実行日 And _
                            '   G借入金入力(j).日割日数 <> 0 And _
                            '    (p借入計画マスタ.利息控除区分 = 2 Or p借入計画マスタ.利息控除区分 = 3) Then
                            '    G借入金入力(j).日割日数 = G借入金入力(j).日割日数 + 1
                            'End If
                            
                            
                            w残高開始年月日 = G借入金入力(l - 1).借入返済年月日
                            w残高終了年月日 = DateAdd("D", -1, G借入金入力(j).借入返済年月日)
                        
                            w利息残高開始年月日 = DateAdd("D", 1, G借入金入力(l - 1).利息計算年月日)
                            w利息残高終了年月日 = G借入金入力(j).利息計算年月日
                            
                               
                            
                            
                            'wx利息残高終了年月日 = w利息残高終了年月日  '2012/08/21
                            
                            'If G借入金入力(j).借入返済年月日 = p借入計画マスタ.最終返済実行日 And _
                            '    (p借入計画マスタ.利息控除区分 = 2 Or p借入計画マスタ.利息控除区分 = 3) Then '2012/08/28
                            '    wx利息残高終了年月日 = DateAdd("D", -1, wx利息残高終了年月日) '2012/08/28
                            'End If                                      '2012/08/28
                            
                            
                            'wx利息残高開始年月日 = DateAdd("D", -G借入金入力(j).日割日数 + 1, wx利息残高終了年月日) '2012/08/21
                            
                            
                            w借入残 = G借入金入力(j).融資残高 + G借入金入力(j).元金
                        
                            If G借入金入力(l - 1).借入返済年月日 = p借入計画マスタ.実行日 Then
                                w利息残高開始年月日 = G借入金入力(l - 1).利息計算年月日
                            Else
                                w利息残高開始年月日 = DateAdd("D", 1, G借入金入力(l - 1).利息計算年月日)
                            End If
                            
                            
                            
                            
                            
                            
                            '***中間利払いの時　開始年月日　終了年月日　を　1日減ずる
                            p借入計画マスタ.借入番号 = p借入計画マスタ.借入番号 'テスト用
                            If p借入計画マスタ.利息控除区分 = 4 Then              '2016/09/29
                             If p借入計画マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then               '2016/09/28
                                    '**利息先払いの時
                                If Format(p借入計画マスタ.実行日, "yyyy/mm/dd") = Format(w利息残高開始年月日, "yyyy/mm/dd") Then  '2016/09/27
                                Else                                    '2016/09/27
                                        w利息残高開始年月日 = DateAdd("d", -1, w利息残高開始年月日)       '2016/09/29
                                        w利息残高終了年月日 = DateAdd("d", -1, w利息残高終了年月日)     '2016/09/29
                                End If                                              '2016/09/29
                             Else                                                    '2016/09/29
                                '**利息後払いの時
                                If Format(p借入計画マスタ.実行日, "yyyy/mm/dd") = Format(w利息残高開始年月日, "yyyy/mm/dd") Then  '2016/09/27
                                    
                                    w利息残高終了年月日 = DateAdd("d", -1, w利息残高終了年月日)     '2016/09/29
                                Else
                                    w利息残高開始年月日 = DateAdd("d", -1, w利息残高開始年月日)     '2016/09/29
                                    w利息残高終了年月日 = DateAdd("d", -1, w利息残高終了年月日)     '2016/09/29
                                End If
                             End If                                              '2016/09/29
                            End If
                       
                            
                            
                            
                            
                            wx利息残高終了年月日 = w利息残高終了年月日  '2016/10/04
                            
                            If G借入金入力(j).借入返済年月日 = p借入計画マスタ.最終返済実行日 And _
                                (p借入計画マスタ.利息控除区分 = 2 Or p借入計画マスタ.利息控除区分 = 3) Then '2016/10/04
                                wx利息残高終了年月日 = DateAdd("D", -1, wx利息残高終了年月日) '2016/10/04
                            End If                                      '2016/10/04
                            
                            
                            wx利息残高開始年月日 = DateAdd("D", -G借入金入力(j).日割日数 + 1, wx利息残高終了年月日) '2016/10/04
                            
                            
                            wx日割日数 = G借入金入力(j).日割日数        '2012/08/21
                            
                            
                            If G借入金入力(j).日割日数 = 0 Then
                                G借入金入力(j).日割日数 = DateDiff("D", w利息残高開始年月日, w利息残高終了年月日) + 1
                                wx日割日数 = DateDiff("D", wx利息残高開始年月日, wx利息残高終了年月日) + 1  '2012/08/21
                                
                            End If
                            
                            GoTo STA3
                        
                        Else
                            l = l - 1
                            GoTo STA2
                        End If
                        
                    Else
                        GoTo 対象外平均残高
                    End If
                End If
                
STA3:

    
                '***平均残高計算
                w対象年月日 = w残高開始年月日
                
STA5:
                w次回年月日 = MRB010_次回年月日算出(w対象年月日)
                
                If w残高終了年月日 > w次回年月日 Then
                       
                    w日数 = DateDiff("D", w対象年月日, w次回年月日) + 1
                    
                    
                Else
                    '***最終回
                    
                    w日数 = DateDiff("D", w対象年月日, w残高終了年月日) + 1
                    
                    
                End If
                
                '*平均残高集計
                For k = 1 To w最終回
                    If w次回年月日 > w年月(k - 1) And w次回年月日 <= w年月(k) Then
                        w平均残高(k) = w平均残高(k) + Round(w借入残 * w日数)
                        Exit For
                    End If
                Next
                
                If w残高終了年月日 > w次回年月日 Then
                    w対象年月日 = DateAdd("D", 1, w次回年月日)
                    GoTo STA5
                End If
                      
                  
'対象外平均残高:


                '***利息平均残高計算
                
                p借入計画マスタ.借入番号 = p借入計画マスタ.借入番号 'テスト用
                
                w対象年月日 = w利息残高開始年月日
                wx対象年月日 = wx利息残高開始年月日                 '2012/08/21
                
                If G借入金入力(j).借入返済年月日 = p借入計画マスタ.実行日 Then  '2016/10/11
                    w対象年月日 = p借入計画マスタ.実行日                        '2016/10/11
                End If                                                          '2016/10/11
               
                
                
                
STA6:
                w次回年月日 = MRB010_次回年月日算出(w対象年月日)
                wx次回年月日 = MRB010_次回年月日算出(wx対象年月日)
                
                'If G借入金入力(j).借入返済年月日 = p借入計画マスタ.実行日 Then  '2012/08/20
                '    w対象年月日 = p借入計画マスタ.実行日                        '2012/08/20
                'End If                                                          '2012/08/20
                
                If w利息残高終了年月日 > w次回年月日 Then
                    If w対象年月日 = p借入計画マスタ.実行日 And _
                        (p借入計画マスタ.利息控除区分 = 1 Or p借入計画マスタ.利息控除区分 = 3) Then
                        w日数 = DateDiff("D", w対象年月日, w次回年月日)         '2012/07/13
                        p日数 = w日数 + 1                                               '2012/07/13
                    Else                                                                '2012/07/13
                    
                        w日数 = DateDiff("D", w対象年月日, w次回年月日) + 1
                        p日数 = w日数                                                   '2012/07/13
                    End If                                                              '2012/07/13
                    
                    
                    
                Else
                    '***最終回
                    
                    w日数 = DateDiff("D", w対象年月日, w利息残高終了年月日) + 1
                    p日数 = w日数
                    
                    
                    
                End If
                
                '***利息按分日数
                If wx利息残高終了年月日 > wx次回年月日 Then         '2012/08/21
                    wx日数 = DateDiff("D", wx対象年月日, wx次回年月日) + 1  '2012/08/21
                Else                                                '2012/08/21
                    wx日数 = DateDiff("D", wx対象年月日, wx利息残高終了年月日) + 1  '2012/08/21
                End If                                              '2012/08/21
                
                
                
                '*利息計算平均残高集計
                w日数残 = w日数残 + p日数   '2012/07/13
                wx日数残 = wx日数残 + wx日数 '2012/08/21
                
                    
                'If w日数残 >= G借入金入力(j).日割日数 Or w利息残高終了年月日 <= w次回年月日 Then
                '    w利息基礎額 = G借入金入力(j).利息額 - w利息額残高
                'Else
                '    w利息基礎額 = Fix(G借入金入力(j).利息額 * w日数 / G借入金入力(j).日割日数)  '2012/03/23
                'End If
                
                If wx日数残 >= wx日割日数 Or wx利息残高終了年月日 <= wx次回年月日 Then  '2012/08/21
                    w利息基礎額 = G借入金入力(j).利息額 - w利息額残高       '2012/08/21
                Else                                                        '2012/08/21
                    w利息基礎額 = Fix(G借入金入力(j).利息額 * wx日数 / wx日割日数)  '2012/08/21
                End If                                                      '2012/08/21
                
                    
                w利息額残高 = w利息額残高 + w利息基礎額
                
                p借入計画マスタ.借入番号 = p借入計画マスタ.借入番号 'テスト用
                
                For k = 1 To w最終回
                    If w次回年月日 > w年月(k - 1) And w次回年月日 <= w年月(k) Then
                        w日数 = p日数                           '2012/07/13
                        w利息計算平均残高(k) = w利息計算平均残高(k) + Round(w借入残 * w日数)
                    
                        'w平均利息基礎額(k) = w平均利息基礎額(k) + w利息基礎額  2012/08/21
                        Exit For
                    End If
                Next
                
                
                For k = 1 To w最終回
                    If wx次回年月日 > w年月(k - 1) And wx次回年月日 <= w年月(k) Then    '2012/08/21
                        w日数 = p日数                           '2012/07/13
                        'w利息計算平均残高(k) = w利息計算平均残高(k) + Round(w借入残 * w日数)
                    
                        w平均利息基礎額(k) = w平均利息基礎額(k) + w利息基礎額  '2012/08/21
                        Exit For
                    End If
                Next
                
                
                If w利息残高終了年月日 > w次回年月日 Then
                    w対象年月日 = DateAdd("D", 1, w次回年月日)
                    wx対象年月日 = DateAdd("D", 1, wx次回年月日)    '2012/08/21
                    GoTo STA6
                End If
  
                
                
対象外平均残高:

対象外利息平均残高:

            
            
            Next
            
            
              
             For k = 1 To wcnt                  '11/06/15 V200
                w平均残高日数(k) = DateDiff("D", w年月(k - 1), w年月(k)) '11/06/15 V200
             Next                               '11/06/15 V200
             
             
             For k = 1 To wcnt
                w平均残高合計 = w平均残高合計 + w平均残高(k)
                
                w利息計算平均残高合計 = w利息計算平均残高合計 + w利息計算平均残高(k)    '11/06/15 V200
                w平均残高日数合計 = w平均残高日数合計 + w平均残高日数(k)    '11/06/15 V200
                w平均利息基礎額合計 = w平均利息基礎額合計 + w平均利息基礎額(k)  '11/06/15 V200
             Next
              
             '***** 平均残高等による　DCDA010_借入残高推移表結果２　の　更新処理
             If w平均残高合計 <> 0 Or _
                w利息計算平均残高合計 <> 0 Or _
                w平均利息基礎額合計 <> 0 Then
                
                wstr2 = ""
                wstr2 = wstr2 + "Select * From DCDA010_借入残高推移表結果２"
                wstr2 = wstr2 + " Where 借入番号='" & p借入計画マスタ.借入番号 & "'"
                Call AdoRecordsetOpen(GDb, wRs2, wstr2)
                If wRs2.eof Then
                    wRs2.AddNew
                    wRs2("借入番号") = p借入計画マスタ.借入番号
                End If
                
                    wRs2("平均残高合計") = w平均残高合計
                    
                    wRs2("利息計算平均残高合計") = w利息計算平均残高合計    '11/06/15 V200
                    wRs2("平均残高日数合計") = w平均残高日数合計            '11/06/15 V200
                    wRs2("平均利息基礎額合計") = w平均利息基礎額合計        '11/06/15 V200
                    
                    For k = 1 To wcnt
                        wRs2("平均残高_" + CStr(Format(k, "00"))) = w平均残高(k)    '11/06/15 V200
                        
                        wRs2("利息計算平均残高_" + CStr(Format(k, "00"))) = w利息計算平均残高(k)
                        wRs2("平均残高日数_" + CStr(Format(k, "00"))) = w平均残高日数(k)
                        wRs2("平均利息基礎額_" + CStr(Format(k, "00"))) = w平均利息基礎額(k)
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
MRB010_手入力平均利率_ERR:
    pERR_MES = pPROGRAM_ID + "/ MRB010_手入力平均利率() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)
Resume
    End
    

End Sub


'------------------------------------------------
' MRB010_次回年月日算出
'------------------------------------------------
Public Function MRB010_次回年月日算出(p対象年月日 As Variant) As Variant
                                      
 '
   
    
    Dim w年 As Integer
    Dim w月 As Integer
    Dim w日 As Integer
    Dim w閏年 As Integer
    
     
    
'
    On Error GoTo MRB010_次回年月日算出_ERR
'
    
    w年 = Format(p対象年月日, "yyyy")
    w月 = Format(p対象年月日, "mm")
    w日 = Format(p対象年月日, "dd")
    
    If G基本情報.決算締日 < w日 Then
        w月 = w月 + 1
        If w月 > 12 Then
            w月 = 1
            w年 = w年 + 1
        End If
        
    End If
    
    w日 = G基本情報.決算締日
    
    w閏年 = w年 Mod 4
    
    If w日 >= 29 Then
        If w月 = 2 Then
            If w閏年 = 0 Then
                w日 = 29
            Else
                w日 = 28
            End If
        End If
    End If
    
    If w日 = 31 And (w月 = 4 Or w月 = 6 Or w月 = 9 Or w月 = 11) Then
        w日 = 30
    End If
    
    
    MRB010_次回年月日算出 = Format(CStr(w年) & "/" & CStr(w月) & "/" & CStr(w日))
    
    
    
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MRB010_次回年月日算出_ERR:
    pERR_MES = pPROGRAM_ID + "/ MRB010_次回年月日算出() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

 
