Attribute VB_Name = "MAA200_時価評価"
Option Explicit
'
Private Const pPROGRAM_ID As String = "MAA200_時価評価"

'----------------< 長期プライムレート >-------------------------
Type MAA200_基準金利レート
    基準金利区分 As String
    年月日() As Date
    レート() As Double
End Type
Type MAA200_選択基準金利レート
    年月日 As Date
    レート As Double
End Type

'------------------------------------------------
' MAA200_基準金利レート設定
'------------------------------------------------
Public Sub MAA200_基準金利レート設定()
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    Dim j As Integer, k As Integer
    Dim wiCnt As Integer
'
    On Error GoTo MAA200_基準金利レート設定_ERR
'
    wstr = ""
    wstr = wstr & "SELECT Count(*) As カウント From DAAA116_基準金利"
    wstr = wstr & " Where 取消フラグ=0"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs("カウント") = 0 Then
            wRs.Close
            Set wRs = Nothing
            
            ReDim G基準金利(0)
            ReDim G基準金利(0).年月日(0)
            ReDim G基準金利(0).レート(0)
                    
            Exit Sub
        End If
    
        ReDim G基準金利(wRs("カウント") - 1)
        
    wRs.Close
    Set wRs = Nothing
    
    wiCnt = 0
    
    wstr = ""
    wstr = wstr & "SELECT * From DAAA116_基準金利"
    wstr = wstr & " Where 取消フラグ=0"
    wstr = wstr & " Order by 基準金利区分"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    Do Until wRs.eof
        G基準金利(wiCnt).基準金利区分 = wRs("基準金利区分")
        wiCnt = wiCnt + 1
        
        wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
'
    For j = 0 To UBound(G基準金利)
        wstr = ""
        wstr = wstr + "Select Count(*) As カウント From DBDA010_借入金長期プライムレート"
        wstr = wstr & " Where 基準金利区分='" & G基準金利(j).基準金利区分 & "'"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
            If wRs("カウント") = 0 Then
                ReDim Preserve G基準金利(j).年月日(0)
                ReDim Preserve G基準金利(j).レート(0)
            Else
                ReDim Preserve G基準金利(j).年月日(wRs("カウント") - 1)
                ReDim Preserve G基準金利(j).レート(wRs("カウント") - 1)
            End If
        wRs.Close
        Set wRs = Nothing
    
        wstr = ""
        wstr = wstr + "Select * From DBDA010_借入金長期プライムレート"
        wstr = wstr & " Where 基準金利区分='" & G基準金利(j).基準金利区分 & "'"
        
        wstr = wstr & " Order by 年月日 Asc" '20170411 ADD M.Mino
        
        Call AdoRecordsetOpen(GDb, wRs, wstr)
            k = -1
            
            Do Until wRs.eof
                k = k + 1
                    
                G基準金利(j).年月日(k) = wRs("年月日")
                G基準金利(j).レート(k) = wRs("長期プライムレート")
                 
                wRs.MoveNext
            Loop
        wRs.Close
        Set wRs = Nothing
    Next
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA200_基準金利レート設定_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA200_基準金利レート設定() でエラー" + vbCrLf + vbCrLf + _
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
' MAA200_基準金利レートRead
'------------------------------------------------
Public Function MAA200_基準金利レートRead(p年月日 As Variant, p基準金利区分 As String) As MAA200_選択基準金利レート
'
    Dim j As Integer, k As Integer
    Dim wdate As Date
'
    On Error GoTo MAA200_基準金利レートRead_ERR
'
    p年月日 = Format(p年月日, "yyyy/mm/dd")
    
    MAA200_基準金利レートRead.レート = 0
'
    For k = 0 To UBound(G基準金利)
        If G基準金利(k).基準金利区分 = p基準金利区分 Then
            For j = 0 To UBound(G基準金利(k).年月日)
                wdate = Format(G基準金利(k).年月日(j), "yyyy/mm/dd")
                If j = 0 Then
                    MAA200_基準金利レートRead.年月日 = G基準金利(k).年月日(j)
                    MAA200_基準金利レートRead.レート = P8.FCDblRD5(G基準金利(k).レート(j))
                    If wdate >= p年月日 Then
                        Exit For
                    End If
                Else
                    If wdate > p年月日 Then
                        Exit For
                    ElseIf wdate = p年月日 Then
                        MAA200_基準金利レートRead.年月日 = G基準金利(k).年月日(j)
                        MAA200_基準金利レートRead.レート = P8.FCDblRD5(G基準金利(k).レート(j))
        
                        Exit For
                    End If
                    MAA200_基準金利レートRead.年月日 = G基準金利(k).年月日(j)
                    MAA200_基準金利レートRead.レート = P8.FCDblRD5(G基準金利(k).レート(j))
                End If
            Next
        End If
    Next
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA200_基準金利レートRead_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA200_基準金利レートRead() でエラー" + vbCrLf + vbCrLf + _
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
' MAA200_時価評価適用金利作成
'------------------------------------------------
Public Function MAA200_時価評価適用金利作成(p決算日 As Date, p基準金利区分 As String, p金利種別 As String) As Boolean
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    
    Dim w借入金 As MAA910_借入金
    Dim w基準金利レート As MAA200_選択基準金利レート
    
    Dim wiCnt As Integer
    Dim j As Integer
    Dim FLG_Data As Boolean
    Dim wdPRM As Double, wd適用金利 As Double
    Dim ws銀行番号 As String
'
    On Error GoTo MAA200_時価評価適用金利作成_ERR
'
    MAA200_時価評価適用金利作成 = False
    
    wstr = ""
    wstr = wstr & "Select K.* From DBDA010_借入金 As K"
    wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
    wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分"
    'wstr = wstr & " Where K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "固定金利")) DEL 20170501 M.Mino
    
    'ADD 20170501 M.Mino
    If p金利種別 = "固定金利" Then
        wstr = wstr & " Where K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "固定金利"))
    Else
        wstr = wstr & " Where K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利"))
    End If
    'ADD END 20170501
    
    wstr = wstr & " and   K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "長期借入金"))
    'wstr = wstr & " and  K.借入金種別区分 = '01'"
    wstr = wstr & " and   K.基準金利区分 = '" & p基準金利区分 & "'"
    wstr = wstr & " and Format(K.実行日,'yyyymmdd') <= '" & Format(p決算日, "yyyymmdd") & "'"
    wstr = wstr & " and Format(K.最終返済実行日,'yyyymmdd') > '" & Format(p決算日, "yyyymmdd") & "'"
    wstr = wstr & " and (K.解約実行日 is null "
    wstr = wstr & " or Format(K.解約実行日,'yyyymmdd') > '" & Format(p決算日, "yyyymmdd") & "'"
    wstr = wstr & " )"
    wstr = wstr & " And K.手入力区分 <> 2"
    wstr = wstr & " And S.社債フラグ=0"
    '16/03/26 利子補給に伴う変更
    wstr = wstr & " And S.利子補給金フラグ=0"
    wstr = wstr & " And K.取消フラグ = 0"
    'wstr = wstr & " Order by 借入金種別区分,銀行番号,実行日"
    wstr = wstr & " Order by K.銀行番号,K.実行日"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.eof Then
        wRs.Close
        Set wRs = Nothing
        
        ReDim G適用金利(0)
        
        Exit Function
    End If
    
    ReDim G適用金利(wRs.RecordCount)
    wiCnt = 0
        
        Do Until wRs.eof
                            
        w借入金 = MBD010_借入データセット(wRs)
        
        '解約Check
        FLG_Data = True
        If P8.FCStr(w借入金.解約実行日) <> "" _
        And Format(w借入金.解約実行日, "yyyymmdd") < Format(p決算日, "yyyymmdd") Then
            FLG_Data = False
        End If
        
        If FLG_Data = True Then
            
            wiCnt = wiCnt + 1
        
            G適用金利(wiCnt).借入番号 = w借入金.借入番号
            G適用金利(wiCnt).銀行番号 = w借入金.銀行番号
            G適用金利(wiCnt).基準金利区分 = w借入金.基準金利区分
            G適用金利(wiCnt).実行日 = w借入金.実行日
            G適用金利(wiCnt).融資金額 = w借入金.融資金額
            G適用金利(wiCnt).利率 = w借入金.利率
            G適用金利(wiCnt).最終返済実行日 = w借入金.最終返済実行日
            G適用金利(wiCnt).決算日 = p決算日
            
            '借入時長期プライムレート
            w基準金利レート = MAA200_基準金利レートRead(w借入金.実行日, w借入金.基準金利区分)
            G適用金利(wiCnt).借入時長プラ = w基準金利レート.レート
            '借入時プレミアム
            G適用金利(wiCnt).借入時PRM = P8.FCDblRD5(w借入金.利率) - G適用金利(wiCnt).借入時長プラ
            
            '決算時融資残高
            G適用金利(wiCnt).決算時融資残高 = 0
            If w借入金.手入力区分 = P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
                G適用金利(wiCnt).決算時融資残高 = MBD010_借入金標準入力残高(w借入金, "", 1, P8.FCDate(G適用金利(wiCnt).決算日), XMXA020_区分("借入金管理区分", "決算用"))
            Else
                Call MBD010_借入金入力明細Read(w借入金)
                G適用金利(wiCnt).決算時融資残高 = MBD010_借入金手入力残高(w借入金, 1, P8.FCDate(G適用金利(wiCnt).決算日))
            End If
            
            '決算時長期プライムレート
            w基準金利レート = MAA200_基準金利レートRead(G適用金利(wiCnt).決算日, w借入金.基準金利区分)
            G適用金利(wiCnt).決算時長プラ = w基準金利レート.レート
            
        End If
        
        wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
'
    '金融機関毎に最新値 時価評価適用プレミアムを取得後、決算時時価評価適用金利を算出
    ws銀行番号 = ""
    wdPRM = 0
    For j = 1 To UBound(G適用金利)
            wd適用金利 = 0
            
            If G適用金利(j).銀行番号 <> ws銀行番号 Then
                ws銀行番号 = G適用金利(j).銀行番号
                wdPRM = MAA200_時価評価適用PRM(ws銀行番号)
            End If
    
            If G適用金利(j).銀行番号 = ws銀行番号 Then
                G適用金利(j).決算時適用PRM = wdPRM
                G適用金利(j).決算時適用金利 = G適用金利(j).決算時長プラ + G適用金利(j).決算時適用PRM
            End If
    Next j
    
    MAA200_時価評価適用金利作成 = True
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA200_時価評価適用金利作成_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA200_時価評価適用金利作成() でエラー" + vbCrLf + vbCrLf + _
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
' MAA200_時価評価適用PRM
'------------------------------------------------
Private Function MAA200_時価評価適用PRM(p銀行番号 As String) As Double
'
    Dim k As Integer
    Dim w実行日 As Variant
    Dim FLG_G As Boolean
'
    On Error GoTo MAA200_時価評価適用PRM_ERR
'
    w実行日 = CDate("1900/01/01")
    FLG_G = False
    
    For k = 1 To UBound(G適用金利)
        If G適用金利(k).銀行番号 = p銀行番号 Then
            If G適用金利(k).実行日 > w実行日 Then
                MAA200_時価評価適用PRM = G適用金利(k).借入時PRM
                w実行日 = G適用金利(k).実行日
            End If
        
            FLG_G = True
        End If
    
        If G適用金利(k).銀行番号 <> p銀行番号 _
        And FLG_G = True Then
            Exit For
        End If
    Next k
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA200_時価評価適用PRM_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA200_時価評価適用PRM() でエラー" + vbCrLf + vbCrLf + _
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
' MAA200_時価評価適用金利
'------------------------------------------------
Public Function MAA200_時価評価適用金利(Optional pKubun As String = "") As Boolean
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String

    Dim j As Integer
    Dim wsTbl As String
'
    On Error GoTo MAA200_時価評価適用金利_ERR
'
    MAA200_時価評価適用金利 = False
    
    If UBound(G適用金利) < 1 Then
        Exit Function
    End If
    
    If G適用金利(1).借入番号 = "" Then
        Exit Function
    End If
    
    wsTbl = "DCDA010_借入金時価評価適用金利"
    If pKubun = "前期末" Then
        wsTbl = "DCDA010_借入金時価評価適用金利前期末"
    End If

    wstr = ""
    wstr = wstr + "Select * From " & wsTbl
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        For j = 1 To UBound(G適用金利)
        
            wRs.AddNew
            
            wRs("借入番号") = G適用金利(j).借入番号
            wRs("銀行番号") = G適用金利(j).銀行番号
            wRs("基準金利区分") = G適用金利(j).基準金利区分
            wRs("実行日") = G適用金利(j).実行日
            wRs("融資金額") = G適用金利(j).融資金額
            wRs("利率") = G適用金利(j).利率
            wRs("最終返済実行日") = G適用金利(j).最終返済実行日
            wRs("決算年月日") = G適用金利(j).決算日
            wRs("決算時融資残高") = G適用金利(j).決算時融資残高
            wRs("借入時長期プライムレート") = G適用金利(j).借入時長プラ
            wRs("借入時プレミアム") = G適用金利(j).借入時PRM
            wRs("決算時長期プライムレート") = G適用金利(j).決算時長プラ
            wRs("時価評価適用プレミアム") = G適用金利(j).決算時適用PRM
            wRs("決算時時価評価適用金利") = G適用金利(j).決算時適用金利

            wRs.Update
          
        Next
    wRs.Close
    Set wRs = Nothing
    
    MAA200_時価評価適用金利 = True
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA200_時価評価適用金利_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA200_時価評価適用金利() でエラー" + vbCrLf + vbCrLf + _
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
' MAA200_Get適用金利
'------------------------------------------------
Public Function MAA200_Get適用金利(p借入番号 As String) As MAA910_適用金利
'
    Dim j As Integer
'
    On Error GoTo MAA200_Get適用金利_ERR
'
    For j = 1 To UBound(G適用金利)
        If G適用金利(j).借入番号 = p借入番号 Then
            MAA200_Get適用金利.借入番号 = G適用金利(j).借入番号
            MAA200_Get適用金利.銀行番号 = G適用金利(j).銀行番号
            MAA200_Get適用金利.基準金利区分 = G適用金利(j).基準金利区分
            MAA200_Get適用金利.実行日 = G適用金利(j).実行日
            MAA200_Get適用金利.融資金額 = G適用金利(j).融資金額
            MAA200_Get適用金利.利率 = G適用金利(j).利率
            MAA200_Get適用金利.最終返済実行日 = G適用金利(j).最終返済実行日
            MAA200_Get適用金利.決算日 = G適用金利(j).決算日
            MAA200_Get適用金利.決算時融資残高 = G適用金利(j).決算時融資残高
            MAA200_Get適用金利.借入時長プラ = G適用金利(j).借入時長プラ
            MAA200_Get適用金利.借入時PRM = G適用金利(j).借入時PRM
            MAA200_Get適用金利.決算時長プラ = G適用金利(j).決算時長プラ
            MAA200_Get適用金利.決算時適用PRM = G適用金利(j).決算時適用PRM
            MAA200_Get適用金利.決算時適用金利 = G適用金利(j).決算時適用金利
            
            Exit For
        End If
    Next
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA200_Get適用金利_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA200_Get適用金利() でエラー" + vbCrLf + vbCrLf + _
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
' MAA200_時価評価明細
'------------------------------------------------
Public Function MAA200_時価評価明細() As Double
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String

    Dim j As Integer
    Dim wd01 As Double
    Dim FFLG As Boolean
'
    On Error GoTo MAA200_時価評価明細_ERR
'
    FFLG = False
    wd01 = 0
    
    wstr = ""
    wstr = wstr + "Select * From DCDA010_借入金時価評価明細"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        For j = 1 To UBound(G時価明細)
            If G時価明細(j).借入番号 <> "" Then
                wRs.AddNew
                            
                wRs("借入番号") = G時価明細(j).借入番号
                wRs("決算年月日") = G時価明細(j).決算日
                wRs("時価評価金利") = G時価明細(j).適用金利
                wRs("返済年月日") = G時価明細(j).返済年月日
                wRs("利息計算年月日") = G時価明細(j).利息計算年月日
                wRs("元金額") = G時価明細(j).元金額
                wRs("利息額") = G時価明細(j).利息額
                wRs("返済金額") = G時価明細(j).返済金額
                wRs("融資残高") = G時価明細(j).融資残高
                wRs("日割日数") = G時価明細(j).日割日数
                wRs("指数") = G時価明細(j).指数
                wRs("分母") = G時価明細(j).分母
                wRs("現価係数") = G時価明細(j).現価係数
                wRs("現在価値") = G時価明細(j).現在価値
                
                If FFLG = False Then
                    wd01 = G時価明細(j).元金額 + G時価明細(j).融資残高
                    FFLG = True
                End If
                
                wRs.Update
            End If
        Next
    
    wRs.Close
    Set wRs = Nothing
    
    MAA200_時価評価明細 = wd01
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA200_時価評価明細_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA200_時価評価明細() でエラー" + vbCrLf + vbCrLf + _
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
' MAA200_時価評価明細作成
'------------------------------------------------
Public Function MAA200_時価評価明細作成(p借入計画マスタ As MAA910_借入金, p決算日 As Date) As MAA910_時価評価一覧
'
    Dim w適用金利 As MAA910_適用金利
    
    Dim j As Integer
    Dim wiCnt As Integer
    Dim w解約実行日 As Variant      '07/02/21 V180
    Dim w返済回数 As Integer        '10/02/27
    Dim w来期決算日 As Date
'
    On Error GoTo MAA200_時価評価明細作成_ERR
'
    ReDim G時価明細(UBound(G借入金テーブル))
    
    w来期決算日 = C年月日.GetDate("設定", DateAdd("yyyy", 1, p決算日), G基本情報.決算締日)
'
    'G適用金利から該当する時価評価適用金利取得
    w適用金利 = MAA200_Get適用金利(p借入計画マスタ.借入番号)
    If w適用金利.決算日 = "" Then
        Exit Function
    End If

    wiCnt = 0
    For j = 1 To UBound(G借入金テーブル)
        If G借入金テーブル(j).実際年月日 > p決算日 Then
            If G借入金テーブル(j).元金額 <> 0 Or G借入金テーブル(j).利息額 <> 0 _
            Or (G借入金テーブル(j).融資残高 <> 0 _
                And p借入計画マスタ.実行日 = G借入金テーブル(j).実際年月日) _
                Or G借入金テーブル(j).保証料 <> 0 Or G借入金テーブル(j).手数料 <> 0 _
                Or Format(w解約実行日, "yyyymmdd") = Format(G借入金テーブル(j).実際年月日, "yyyymmdd") Then '10/06/16 V195
               
                wiCnt = wiCnt + 1
                
                G時価明細(wiCnt).借入番号 = G借入金テーブル(j).借入番号
                G時価明細(wiCnt).返済年月日 = G借入金テーブル(j).実際年月日
                G時価明細(wiCnt).利息計算年月日 = G借入金テーブル(j).利息計算年月日   '10/01/04
                G時価明細(wiCnt).利息額 = G借入金テーブル(j).利息額
                
                ' 07/02/21 V180
                w解約実行日 = p借入計画マスタ.解約実行日
                If Format(w解約実行日, "yyyymmdd") = Format(G借入金テーブル(j).実際年月日, "yyyymmdd") Then
                    G時価明細(wiCnt).返済金額 = G借入金テーブル(j).融資残高 + G借入金テーブル(j).利息額
                    G時価明細(wiCnt).元金額 = G借入金テーブル(j).融資残高
                    G時価明細(wiCnt).融資残高 = 0
                Else
                    G時価明細(wiCnt).返済金額 = G借入金テーブル(j).返済金額
                    G時価明細(wiCnt).元金額 = G借入金テーブル(j).元金額
                    G時価明細(wiCnt).融資残高 = G借入金テーブル(j).融資残高
                End If
                    
                '時価評価 2015/06/01 V430
                G時価明細(wiCnt).決算日 = w適用金利.決算日
                G時価明細(wiCnt).適用金利 = w適用金利.決算時適用金利
            
                GDate1 = DateAdd("d", 1, CDate(w適用金利.決算日))
                G時価明細(wiCnt).日割日数 = DateDiff("d", GDate1, CDate(G借入金テーブル(j).実際年月日)) + 1
                G時価明細(wiCnt).指数 = G時価明細(wiCnt).日割日数 / 365
                G時価明細(wiCnt).分母 = (1 + w適用金利.決算時適用金利 / 100) ^ G時価明細(wiCnt).指数
'                G時価明細(wiCnt).現価係数 = P8.FCDiv(1, G時価明細(wiCnt).分母)
'                G時価明細(wiCnt).指数 = P8.FCDblRD6(G時価明細(wiCnt).日割日数 / 365)
'                G時価明細(wiCnt).分母 = P8.FCDblRD6((1 + w適用金利.決算時適用金利 / 100) ^ G時価明細(wiCnt).指数)
                G時価明細(wiCnt).現価係数 = P8.FCDblRD6(P8.FCDiv(1, G時価明細(wiCnt).分母))
                'G時価明細(wiCnt).現在価値 = P8.FRound(P8.FCDiv(G時価明細(wiCnt).返済金額, G時価明細(wiCnt).分母), 0)
                G時価明細(wiCnt).現在価値 = P8.FRound(G時価明細(wiCnt).返済金額 * G時価明細(wiCnt).現価係数, 0)
                
                '1年以内 1年超 集計
                If w来期決算日 >= G時価明細(wiCnt).返済年月日 Then
                    MAA200_時価評価明細作成.年内元金額 = MAA200_時価評価明細作成.年内元金額 + G時価明細(wiCnt).元金額
                    MAA200_時価評価明細作成.年内現在価値 = MAA200_時価評価明細作成.年内現在価値 + G時価明細(wiCnt).現在価値
                ElseIf w来期決算日 < G時価明細(wiCnt).返済年月日 Then
                    MAA200_時価評価明細作成.年超元金額 = MAA200_時価評価明細作成.年超元金額 + G時価明細(wiCnt).元金額
                    MAA200_時価評価明細作成.年超現在価値 = MAA200_時価評価明細作成.年超現在価値 + G時価明細(wiCnt).現在価値
                End If
                
            End If
        End If
    Next
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA200_時価評価明細作成_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA200_時価評価明細作成() でエラー" + vbCrLf + vbCrLf + _
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
' MAA200_時価評価明細作成_入力登録
'------------------------------------------------
Public Function MAA200_時価評価明細作成_入力登録(p借入計画マスタ As MAA910_借入金, p決算日 As Date) As MAA910_時価評価一覧
'
    Dim j As Integer
    Dim wiCnt As Integer
    Dim w適用金利 As MAA910_適用金利
    Dim w来期決算日 As Date
'
    On Error GoTo MAA200_時価評価明細作成_入力登録_ERR
'
    ReDim G時価明細(UBound(G借入金テーブル))
    
    w来期決算日 = C年月日.GetDate("設定", DateAdd("yyyy", 1, p決算日), G基本情報.決算締日)
'
    'G適用金利から該当する時価評価適用金利取得
    w適用金利 = MAA200_Get適用金利(p借入計画マスタ.借入番号)
    If w適用金利.決算日 = "" Then
        Exit Function
    End If
    
    wiCnt = 0
    For j = 1 To UBound(G借入金入力)
        If G借入金入力(j).借入返済年月日 > p決算日 Then
            wiCnt = wiCnt + 1
            
            G時価明細(wiCnt).借入番号 = p借入計画マスタ.借入番号
            G時価明細(wiCnt).返済年月日 = G借入金入力(j).借入返済年月日
            G時価明細(wiCnt).利息計算年月日 = G借入金入力(j).利息計算年月日
            
            G時価明細(wiCnt).元金額 = G借入金入力(j).元金
            G時価明細(wiCnt).利息額 = G借入金入力(j).利息額
            G時価明細(wiCnt).返済金額 = G借入金入力(j).返済金額
            G時価明細(wiCnt).融資残高 = G借入金入力(j).融資残高
            
            '時価評価 2015/06/01 V430
            G時価明細(wiCnt).決算日 = w適用金利.決算日
            G時価明細(wiCnt).適用金利 = w適用金利.決算時適用金利
        
            GDate1 = DateAdd("d", 1, CDate(w適用金利.決算日))
            G時価明細(wiCnt).日割日数 = DateDiff("d", GDate1, G借入金入力(j).借入返済年月日) + 1
            G時価明細(wiCnt).指数 = P8.FCDblRD6(G時価明細(wiCnt).日割日数 / 365)
            G時価明細(wiCnt).分母 = P8.FCDblRD6((1 + w適用金利.決算時適用金利 / 100) ^ G時価明細(wiCnt).指数)
            G時価明細(wiCnt).現価係数 = P8.FCDblRD6(P8.FCDiv(1, G時価明細(wiCnt).分母))
            'G時価明細(wiCnt).現在価値 = P8.FRound(P8.FCDiv(G時価明細(wiCnt).返済金額, G時価明細(wiCnt).分母), 0)
            G時価明細(wiCnt).現在価値 = P8.FRound(G時価明細(wiCnt).返済金額 * G時価明細(wiCnt).現価係数, 0)
        
            '1年以内 1年超 集計
            If w来期決算日 >= G時価明細(wiCnt).返済年月日 Then
                MAA200_時価評価明細作成_入力登録.年内元金額 = MAA200_時価評価明細作成_入力登録.年内元金額 + G時価明細(wiCnt).元金額
                MAA200_時価評価明細作成_入力登録.年内現在価値 = MAA200_時価評価明細作成_入力登録.年内現在価値 + G時価明細(wiCnt).現在価値
            ElseIf w来期決算日 < G時価明細(wiCnt).返済年月日 Then
                MAA200_時価評価明細作成_入力登録.年超元金額 = MAA200_時価評価明細作成_入力登録.年超元金額 + G時価明細(wiCnt).元金額
                MAA200_時価評価明細作成_入力登録.年超現在価値 = MAA200_時価評価明細作成_入力登録.年超現在価値 + G時価明細(wiCnt).現在価値
            End If
        End If
    Next
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA200_時価評価明細作成_入力登録_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA200_時価評価明細作成_入力登録() でエラー" + vbCrLf + vbCrLf + _
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
' MAA200_時価評価一覧作成
'------------------------------------------------
'Public Function MAA200_時価評価一覧作成(p決算日 As Date, p前期末決算日 As Date, p基準金利区分 As String, pKbn As String) As Boolean
Public Function MAA200_時価評価一覧作成(p決算日 As Date, p前期末決算日 As Date, p基準金利区分 As String, pKbn As String, p金利種別 As String) As Boolean 'UPD 20170501 M.Mino
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    Dim wRet As Boolean
    
    Dim w借入データ As MAA910_借入金
    Dim w時価一覧 As MAA910_時価評価一覧
    
    Dim wiCnt As Integer
    Dim j As Integer
    Dim FLG_Data As Boolean
'
    On Error GoTo MAA200_時価評価一覧作成_ERR
'
    MAA200_時価評価一覧作成 = False
    wRet = False
    
    wstr = ""
    wstr = wstr & "Select K.* From DBDA010_借入金 As K"
    wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
    wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分"
    'ADD 20170501 M.Mino
    If p金利種別 = "固定金利" Then
        wstr = wstr & " Where K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "固定金利"))
    Else
        wstr = wstr & " Where K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利"))
    End If
    'ADD END 20170501
    'wstr = wstr & " Where K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "固定金利")) DEL 20170501 M.Mino
    wstr = wstr & " and   K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "長期借入金"))
    'wstr = wstr & " and  K.借入金種別区分 = '01'" '長期借入金
    wstr = wstr & " and  K.基準金利区分 = '" & p基準金利区分 & "'"
    wstr = wstr & " and Format(K.実行日,'yyyymmdd') <= '" & Format(p決算日, "yyyymmdd") & "'"
    wstr = wstr & " and Format(K.最終返済実行日,'yyyymmdd') >= '" & Format(p決算日, "yyyymmdd") & "'"
    wstr = wstr & " and K.手入力区分 <> 2"
    wstr = wstr & " And S.社債フラグ=0"
    '16/03/26 利子補給に伴う変更
    wstr = wstr & " And S.利子補給金フラグ=0"
    wstr = wstr & " And K.取消フラグ = 0"
    'wstr = wstr & " Order by 借入金種別区分,銀行番号,借入番号"
    wstr = wstr & " Order by K.銀行番号,K.実行日"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.eof Then
        wRs.Close
        Set wRs = Nothing
            
        Exit Function
    End If
        Do Until wRs.eof
        
            w借入データ = MBD010_借入データセット(wRs)
            
            'w時価一覧クリア
            w時価一覧.借入番号 = ""
            w時価一覧.銀行番号 = ""
            w時価一覧.基準金利区分 = ""
            w時価一覧.実行日 = Null
            w時価一覧.融資金額 = 0
            w時価一覧.利率 = 0
            w時価一覧.最終返済実行日 = Null
            w時価一覧.決算日 = Null
            w時価一覧.年内元金額 = 0
            w時価一覧.年内現在価値 = 0
            w時価一覧.年超元金額 = 0
            w時価一覧.年超現在価値 = 0
            
            '解約Check
            FLG_Data = True
            If P8.FCStr(w借入データ.解約実行日) <> "" _
            And Format(w借入データ.解約実行日, "yyyymmdd") < Format(p決算日, "yyyymmdd") Then
                ReDim G時価明細(0)
            Else
                
                If P8.FCDbl(wRs("手入力区分")) = "0" Then
                '標準
                    Call MBD010_借入金テーブル作成("", w借入データ)
                    w時価一覧 = MAA200_時価評価明細作成(w借入データ, p決算日)
                Else
                '入力登録
                    Call MBD010_借入金入力明細Read(w借入データ)
                    w時価一覧 = MAA200_時価評価明細作成_入力登録(w借入データ, p決算日)
                End If
            
            End If
            
            w時価一覧.借入番号 = w借入データ.借入番号
            w時価一覧.銀行番号 = w借入データ.銀行番号
            w時価一覧.基準金利区分 = w借入データ.基準金利区分
            w時価一覧.実行日 = w借入データ.実行日
            w時価一覧.融資金額 = w借入データ.融資金額
            w時価一覧.利率 = w借入データ.利率
            w時価一覧.最終返済実行日 = w借入データ.最終返済実行日
            w時価一覧.決算日 = p決算日
            
            wRet = MAA200_時価評価一覧(p決算日, p前期末決算日, w時価一覧, pKbn)
            If wRet = True Then
                MAA200_時価評価一覧作成 = True
            End If
            
            wRs.MoveNext
        Loop
    
    wRs.Close
    Set wRs = Nothing
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA200_時価評価一覧作成_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA200_時価評価一覧作成() でエラー" + vbCrLf + vbCrLf + _
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
' MAA200_時価評価一覧
'------------------------------------------------
Public Function MAA200_時価評価一覧(p決算日 As Date, p前期末決算日 As Date, p時価一覧 As MAA910_時価評価一覧, pKbn As String) As Boolean
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    Dim wsTbl As String
    Dim j As Integer
'
    On Error GoTo MAA200_時価評価一覧_ERR
'
    MAA200_時価評価一覧 = False
    
    wsTbl = "DCDA010_借入金時価評価"
    If pKbn = "前期末" Then
        wsTbl = "DCDA010_借入金時価評価前期末"
    End If
    
    If p時価一覧.借入番号 = "" Then
        Exit Function
    End If
    
    wstr = ""
    wstr = wstr + "Select * From " & wsTbl
    wstr = wstr + " Where 借入番号='" & p時価一覧.借入番号 & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        wRs.AddNew
        
        wRs("借入番号") = p時価一覧.借入番号
        wRs("銀行番号") = p時価一覧.銀行番号
        wRs("基準金利区分") = "03" 'p時価一覧.基準金利区分
        wRs("実行日") = p時価一覧.実行日
        wRs("融資金額") = p時価一覧.融資金額
        wRs("利率") = p時価一覧.利率
        wRs("最終返済実行日") = p時価一覧.最終返済実行日
        wRs("決算年月日") = p時価一覧.決算日
        
        wRs("合計決算時融資残高") = p時価一覧.年内元金額 + p時価一覧.年超元金額
        wRs("合計時価評価額") = p時価一覧.年内現在価値 + p時価一覧.年超現在価値
        wRs("合計時価損益") = (p時価一覧.年内現在価値 + p時価一覧.年超現在価値) - (p時価一覧.年内元金額 + p時価一覧.年超元金額)
        
        wRs("年以内返済予定元金") = p時価一覧.年内元金額
        wRs("年以内返済予定時価評価額") = p時価一覧.年内現在価値
        wRs("年以内返済予定時価損益") = p時価一覧.年内現在価値 - p時価一覧.年内元金額
        
        wRs("年超返済予定元金") = p時価一覧.年超元金額
        wRs("年超返済予定時価評価額") = p時価一覧.年超現在価値
        wRs("年超返済予定時価損益") = p時価一覧.年超現在価値 - p時価一覧.年超元金額
        
        wRs.Update
    
    wRs.Close
    Set wRs = Nothing
'
    '増減 指定決算日－前期末決算日の為　前期金額に*-1 (sum集計)
    wstr = ""
    wstr = wstr + "Select * From DCDA010_借入金時価評価前期末比較増減"
    wstr = wstr + " Where 借入番号='" & p時価一覧.借入番号 & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        wRs.AddNew
        
        wRs("借入番号") = p時価一覧.借入番号
        wRs("銀行番号") = p時価一覧.銀行番号
        wRs("基準金利区分") = "03" 'p時価一覧.基準金利区分
        wRs("実行日") = p時価一覧.実行日
        wRs("融資金額") = p時価一覧.融資金額
        wRs("利率") = p時価一覧.利率
        wRs("最終返済実行日") = p時価一覧.最終返済実行日
        wRs("決算年月日") = p前期末決算日 'p時価一覧.決算日
        
        wRs("合計決算時融資残高") = p時価一覧.年内元金額 + p時価一覧.年超元金額
        wRs("合計時価評価額") = p時価一覧.年内現在価値 + p時価一覧.年超現在価値
        wRs("合計時価損益") = (p時価一覧.年内現在価値 + p時価一覧.年超現在価値) - (p時価一覧.年内元金額 + p時価一覧.年超元金額)
        
        wRs("年以内返済予定元金") = p時価一覧.年内元金額
        wRs("年以内返済予定時価評価額") = p時価一覧.年内現在価値
        wRs("年以内返済予定時価損益") = p時価一覧.年内現在価値 - p時価一覧.年内元金額
        
        wRs("年超返済予定元金") = p時価一覧.年超元金額
        wRs("年超返済予定時価評価額") = p時価一覧.年超現在価値
        wRs("年超返済予定時価損益") = p時価一覧.年超現在価値 - p時価一覧.年超元金額
        
        If pKbn = "前期末" Then
            wRs("合計決算時融資残高") = wRs("合計決算時融資残高") * -1
            wRs("合計時価評価額") = wRs("合計時価評価額") * -1
            wRs("合計時価損益") = wRs("合計時価損益") * -1
            
            wRs("年以内返済予定元金") = wRs("年以内返済予定元金") * -1
            wRs("年以内返済予定時価評価額") = wRs("年以内返済予定時価評価額") * -1
            wRs("年以内返済予定時価損益") = wRs("年以内返済予定時価損益") * -1
            
            wRs("年超返済予定元金") = wRs("年超返済予定元金") * -1
            wRs("年超返済予定時価評価額") = wRs("年超返済予定時価評価額") * -1
            wRs("年超返済予定時価損益") = wRs("年超返済予定時価損益") * -1
        End If
        
        wRs.Update
    
    wRs.Close
    Set wRs = Nothing
'
    MAA200_時価評価一覧 = True
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA200_時価評価一覧_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA200_時価評価一覧() でエラー" + vbCrLf + vbCrLf + _
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
' MAA200_サブタイトル作成
'------------------------------------------------
Public Function MAA200_サブタイトル作成() As String
'
    Dim j As Integer
    Dim wIndex As Integer
    Dim wdate As Date
    Dim ws01 As String
'
    On Error GoTo MAA200_サブタイトル作成_ERR
'
    wdate = G決算日(0)
    wIndex = 0
    
    If G基本情報.決算サイクル = 3 Then
        For j = 1 To 12
            wdate = DateAdd("m", G基本情報.決算サイクル, wdate)
            wIndex = wIndex + 1
            If Format(G決算日(1), "yyyy/mm") = Format(wdate, "yyyy/mm") Then
                ws01 = Format(wdate, Gfmt年月)
                Exit For
            End If
        Next
        If wIndex = 1 Then
            ws01 = Format(wdate, Gfmt年月) & "第一四半期決算"
        ElseIf wIndex = 2 Then
            ws01 = Format(wdate, Gfmt年月) & "第二四半期決算"
        ElseIf wIndex = 3 Then
            ws01 = Format(wdate, Gfmt年月) & "第三四半期決算"
        Else
            ws01 = Format(wdate, Gfmt年月) & "第四四半期決算"
        End If
    ElseIf G基本情報.決算サイクル = 6 Then
        For j = 1 To 12
            wdate = DateAdd("m", G基本情報.決算サイクル, wdate)
            wIndex = wIndex + 1
            If Format(G決算日(1), "yyyy/mm") = Format(wdate, "yyyy/mm") Then
                ws01 = Format(wdate, Gfmt年月)
                Exit For
            End If
        Next
        If wIndex = 1 Then
            ws01 = Format(wdate, Gfmt年月) & "上期決算"
        ElseIf wIndex = 2 Then
            ws01 = Format(wdate, Gfmt年月) & "下期決算"
        Else
            ws01 = Format(wdate, Gfmt年月) & "決算"
        End If
    Else
        wdate = G決算日(1)
        ws01 = Format(wdate, Gfmt年月) & "決算"
    End If
    
    MAA200_サブタイトル作成 = ws01
    
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA200_サブタイトル作成_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA200_サブタイトル作成() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function


