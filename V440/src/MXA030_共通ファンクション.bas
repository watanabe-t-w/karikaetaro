Attribute VB_Name = "MXA030_共通ファンクション"

Option Explicit
'
Private Const pPROGRAM_ID As String = "MXA030_共通ファンクション"
'
'------------------------------------------------
' MXA030_MCLEAR
'------------------------------------------------
Public Sub MXA030_MCLEAR()
'
    ' =========================================
    '         グローバルテーブル　クリア
    ' =========================================
    ReDim G売上計画テーブル(0)
    ReDim G設備計画テーブル(0)
    ReDim G借入金テーブル(0)
        
    Erase G年度計画.年度()
    Erase G年度計画.年度()
    Erase G年度計画.粗利率()
    Erase G年度計画.粗利率1()
    Erase G年度計画.売上1構成比()
    Erase G年度計画.粗利率2()
    Erase G年度計画.売上2構成比()
    Erase G年度計画.粗利率3()
    Erase G年度計画.売上3構成比()
    Erase G年度計画.換算1構成比()
    Erase G年度計画.換算2構成比()
    Erase G年度計画.換算3構成比()
    
    Erase G年度計画.売上回収サイト()
    Erase G年度計画.売上回収1サイト()
    Erase G年度計画.売上回収1サイト1() '4/7/22 V120
    Erase G年度計画.売上回収1サイト2() '4/7/22 V120
    Erase G年度計画.売上回収1サイト3() '4/7/22 V120
    Erase G年度計画.売上回収1構成比1() '4/7/22 V120
    Erase G年度計画.売上回収1構成比2() '4/7/22 V120
    Erase G年度計画.売上回収1構成比3() '4/7/22 V120
    Erase G年度計画.売上回収2サイト()
    Erase G年度計画.売上回収2サイト1() '4/7/22 V120
    Erase G年度計画.売上回収2サイト2() '4/7/22 V120
    Erase G年度計画.売上回収2サイト3() '4/7/22 V120
    Erase G年度計画.売上回収2構成比1() '4/7/22 V120
    Erase G年度計画.売上回収2構成比2() '4/7/22 V120
    Erase G年度計画.売上回収2構成比3() '4/7/22 V120
    Erase G年度計画.売上回収3サイト()
    Erase G年度計画.売上回収3サイト1() '4/7/22 V120
    Erase G年度計画.売上回収3サイト2() '4/7/22 V120
    Erase G年度計画.売上回収3サイト3() '4/7/22 V120
    Erase G年度計画.売上回収3構成比1() '4/7/22 V120
    Erase G年度計画.売上回収3構成比2() '4/7/22 V120
    Erase G年度計画.売上回収3構成比3() '4/7/22 V120
    
    Erase G年度計画.仕入支払サイト()   '05/04/07 V127
    Erase G年度計画.仕入支払1サイト()  '05/04/07 V127
    Erase G年度計画.仕入支払1サイト1() '05/04/07 V127
    Erase G年度計画.仕入支払1サイト2() '05/04/07 V127
    Erase G年度計画.仕入支払1サイト3() '05/04/07 V127
    Erase G年度計画.仕入支払1構成比1() '05/04/07 V127
    Erase G年度計画.仕入支払1構成比2() '05/04/07 V127
    Erase G年度計画.仕入支払1構成比3() '05/04/07 V127
    Erase G年度計画.仕入支払2サイト()  '05/04/07
    Erase G年度計画.仕入支払2サイト1() '05/04/07 V127
    Erase G年度計画.仕入支払2サイト2() '05/04/07 V127
    Erase G年度計画.仕入支払2サイト3() '05/04/07 V127
    Erase G年度計画.仕入支払2構成比1() '05/04/07 V127
    Erase G年度計画.仕入支払2構成比2() '05/04/07 V127
    Erase G年度計画.仕入支払2構成比3() '05/04/07 V127
    Erase G年度計画.仕入支払3サイト3() '05/04/07 V127
    Erase G年度計画.仕入支払3サイト1() '05/04/07 V127
    Erase G年度計画.仕入支払3サイト2() '05/04/07 V127
    Erase G年度計画.仕入支払3サイト3() '05/04/07 V127
    Erase G年度計画.仕入支払3構成比1() '05/04/07 V127
    Erase G年度計画.仕入支払3構成比2() '05/04/07 V127
    Erase G年度計画.売上回収3構成比3() '05/04/07 V127
    
    Erase G年度計画.給与指数()
    Erase G年度計画.賞与指数()
    Erase G年度計画.固定経費指数()
    Erase G年度計画.その他経費1指数()
    Erase G年度計画.保険積立指数()
    Erase G年度計画.給与総額()
    Erase G年度計画.賞与額()
    Erase G年度計画.固定経費()
    Erase G年度計画.変動経費1()
    Erase G年度計画.変動経費2()
    Erase G年度計画.変動経費3()
    Erase G年度計画.その他経費1()
    Erase G年度計画.保険積立()
    Erase G年度計画.営業外収益()
    Erase G年度計画.営業外費用()
    Erase G年度計画.減価償却費()
    Erase G年度計画.支払利息()
'
End Sub

'------------------------------------------------
' MXA030_印字テーブルクリア
'------------------------------------------------
Public Sub MXA030_印字テーブルクリア()
'
    Dim wstr As String
'
'----------< DCAA010_売上計画明細 >-------------------------------------------------
    wstr = "DELETE * FROM DCAA010_売上計画明細"
    GDb.Execute wstr

    DoEvents
'----------< DCAA012_予算基礎表 >---------------------------------------------------
    wstr = "DELETE * FROM DCAA012_予算基礎表"
    GDb.Execute wstr

    DoEvents
'----------< DCAA014_金剛石科目対応表 >---------------------------------------------
    wstr = "DELETE * FROM DCAA014_金剛石科目対応表"
    GDb.Execute wstr

    DoEvents
'----------< DCAA050_損益予実対比表 >-----------------------------------------------
    wstr = "DELETE * FROM DCAA050_損益予実対比表"
    GDb.Execute wstr

    DoEvents
'----------< DCAA060_損益推移表 >---------------------------------------------------
    wstr = "DELETE * FROM DCAA060_損益推移表"
    GDb.Execute wstr

    DoEvents
'----------< DCAA060_損益推移表_Foot >----------------------------------------------
    wstr = "DELETE * FROM DCAA060_損益推移表_Foot"
    GDb.Execute wstr

    DoEvents
'----------< DCAA070_売上詳細卸業分類合計 >-----------------------------------------
    wstr = "DELETE * FROM DCAA070_売上詳細卸業分類合計"
    GDb.Execute wstr

    DoEvents
'----------< DCAA070_売上詳細卸業分類合計2 >----------------------------------------
    wstr = "DELETE * FROM DCAA070_売上詳細卸業分類合計"
    GDb.Execute wstr

    DoEvents
'----------< DCAA070_売上詳細製造業分類合計 >---------------------------------------
    wstr = "DELETE * FROM DCAA070_売上詳細製造業分類合計"
    GDb.Execute wstr

    DoEvents
'----------< DCAA070_売上詳細製造業分類合計2 >--------------------------------------
    wstr = "DELETE * FROM DCAA070_売上詳細製造業分類合計2"
    GDb.Execute wstr

    DoEvents
'----------< DCCA010_設備推移結果 >-------------------------------------------------
    wstr = "DELETE * FROM DCCA010_設備推移結果"
    GDb.Execute wstr

    DoEvents
'----------< DCCA020_設備計画明細 >-------------------------------------------------
    wstr = "DELETE * FROM DCCA020_設備計画明細"
    GDb.Execute wstr

    DoEvents
'----------< DCDA010_借入残高推移表結果 >-------------------------------------------
    wstr = "DELETE * FROM DCDA010_借入残高推移表結果"
    GDb.Execute wstr

    DoEvents
'----------< DCDA020_借入金明細 >---------------------------------------------------
    wstr = "DELETE * FROM DCDA020_借入金明細"
    GDb.Execute wstr

    DoEvents
'----------< DCHA010_Gridワーク >---------------------------------------------------
    wstr = "DELETE * FROM DCHA010_Gridワーク"
    GDb.Execute wstr

    DoEvents
'----------< DCIA010_借入金ワーク >---------------------------------------------------
    wstr = "DELETE * FROM DCIA010_借入金ワーク"
    GDb.Execute wstr

    DoEvents
''----------< DCDA020_貸付金明細 >---------------------------------------------------
'    wstr = "DELETE * FROM DCDA020_貸付金明細"
'    GDb.Execute wstr
'
'    DoEvents
'----------< DCDA030_借入一覧表 >---------------------------------------------------
    wstr = "DELETE * FROM DCDA030_借入一覧表"
    GDb.Execute wstr

    DoEvents
'----------< DCEA010_経営計画支援表 >-----------------------------------------------
    wstr = "DELETE * FROM DCEA010_経営計画支援表"
    GDb.Execute wstr

    DoEvents
'----------< DCEA020_経営計画支援表比較 >-------------------------------------------
    wstr = "DELETE * FROM DCEA020_経営計画支援表比較"
    GDb.Execute wstr

    DoEvents
'----------< DCFA010_決算書登録ワーク >---------------------------------------------
    wstr = "DELETE * FROM DCFA010_決算書登録ワーク"
    GDb.Execute wstr

    DoEvents
'----------< DCFB010_金剛石科目ワーク >---------------------------------------------
    wstr = "DELETE * FROM DCFB010_金剛石科目ワーク"
    GDb.Execute wstr

    DoEvents
'----------< DCFC010_決算書推移表ワーク >-------------------------------------------
    wstr = "DELETE * FROM DCFC010_決算書推移表ワーク"
    GDb.Execute wstr

    DoEvents
'----------< DCFD010_金剛石経費実績推移表ワーク >-----------------------------------
    wstr = "DELETE * FROM DCFD010_金剛石経費実績推移表ワーク"
    GDb.Execute wstr

    DoEvents
'----------< DCGA010_基幹データ調整ワーク >-----------------------------------------
    wstr = "DELETE * FROM DCGA010_基幹データ調整ワーク"
    GDb.Execute wstr

    DoEvents
'----------< DCLA010_ログファイル >-------------------------------------------------
    wstr = "DELETE * FROM DCLA010_ログファイル"
    GDb.Execute wstr

    DoEvents
'----------< DCXA010_帳票作成ワーク >-----------------------------------------------
    wstr = "DELETE * FROM DCXA010_帳票作成ワーク"
    GDb.Execute wstr

    DoEvents
'----------< DCXA020_帳票作成ワーク >-----------------------------------------------
    wstr = "DELETE * FROM DCXA020_帳票作成ワーク"
    GDb.Execute wstr

    DoEvents
'----------< DCXA030_帳票作成ワーク >-----------------------------------------------
    wstr = "DELETE * FROM DCXA030_帳票作成ワーク"
    GDb.Execute wstr

    DoEvents
'----------< DDAA010_売上計画インポート >-------------------------------------------
    wstr = "DELETE * FROM DDAA010_売上計画インポート"
    GDb.Execute wstr

    DoEvents
'----------< DDAA010_売上計画インポートB >------------------------------------------
    wstr = "DELETE * FROM DDAA010_売上計画インポートB"
    GDb.Execute wstr

    DoEvents
'----------< DECA010_科目集計 >-----------------------------------------------------
    wstr = "DELETE * FROM DECA010_科目集計"
    GDb.Execute wstr

    DoEvents
'----------< DECA020_科目集計2 >----------------------------------------------------
    wstr = "DELETE * FROM DECA020_科目集計2"
    GDb.Execute wstr

    DoEvents
'----------< DXAA010_クロス集計結果 >-----------------------------------------------
    wstr = "DELETE * FROM DXAA010_クロス集計結果"
    GDb.Execute wstr

    DoEvents
'----------< DXAA020_科目テーブルデバック >-----------------------------------------
    wstr = "DELETE * FROM DXAA020_科目テーブルデバック"
    GDb.Execute wstr

    DoEvents
'----------< DXAA020_科目テーブルデバック2 >----------------------------------------
    wstr = "DELETE * FROM DXAA020_科目テーブルデバック2"
    GDb.Execute wstr

    DoEvents
'----------< DXAA030_分岐点テスト >-------------------------------------------------
    wstr = "DELETE * FROM DXAA030_分岐点テスト"
    GDb.Execute wstr

    DoEvents
'
End Sub

'------------------------------------------------
' MXA030_DataGridInit
'------------------------------------------------
Public Sub MXA030_DataGridInit(pDataGrid As Control)
'
    On Error GoTo MXA030_DataGridInit_ERR
'
    pDataGrid.AllowRowSizing = False
    pDataGrid.HeadFont.Size = 11
    pDataGrid.HeadFont.Bold = True
    pDataGrid.Font.Size = 11
    pDataGrid.BackColor = C_Yellow
    pDataGrid.ForeColor = RGB(0, 0, 160)
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MXA030_DataGridInit_ERR:
    pERR_MES = pPROGRAM_ID + "/ MXA030_DataGridInit() でエラー" + vbCrLf + vbCrLf + _
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
' MXA030_翌営業年月日計算
'------------------------------------------------
Public Function MXA030_翌営業年月日計算(p年月日 As Date, p支払日 As Integer, p営業日区分 As Integer) As Date
'
    Dim wdate As Date
'
    On Error GoTo MXA030_翌営業年月日計算_ERR
'
    wdate = C年月日.GetDate("設定", p年月日, p支払日)
    
    Call C休日.計算(wdate, p営業日区分)                 ' 07/01/30 V180
    MXA030_翌営業年月日計算 = C休日.次回稼働日
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MXA030_翌営業年月日計算_ERR:
    pERR_MES = pPROGRAM_ID + "/ MXA030_翌営業年月日計算() でエラー" + vbCrLf + vbCrLf + _
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
' MXA030_実行支払年月
'------------------------------------------------
Public Function MXA030_実行支払年月(p年月日 As Variant, p支払日 As Integer, p営業日区分 As Integer, p判断 As String) As Variant ' 07/01/30 V180
'
    Dim wdate As Date
    Dim wDate1 As Date
    Dim wDate2 As Date
    Dim wdd As Integer
'
    On Error GoTo MXA030_実行支払年月_ERR
'
    If IsNull(p年月日) Then
        MXA030_実行支払年月 = Null
    Else
        wdate = C年月日.GetDate("月始", CDate(p年月日))
        wDate1 = wdate
        Select Case p判断
            Case "="
                If Day(p年月日) >= p支払日 Then
                    wdate = DateAdd("m", 1, wdate)
                End If
            
            Case Else
                If Day(p年月日) > p支払日 Then
                    wdate = DateAdd("m", 1, wdate)
                End If
        End Select
         
        MXA030_実行支払年月 = wdate
        
        'If p支払日 = 31 Then                                           ' 07/02/22 V180 修正
        '    wDate2 = DateAdd("m", -1, p年月日) '月末締め
        '    wDate2 = C年月日.GetDate("月末", wDate2)
        '    wDate2 = MXA030_翌営業年月日計算(wDate2, 31, p営業日区分)  ' 07/01/30 V180
        'Else
        '    wDate2 = C年月日.GetDate("月始", wDate1) '月末締め以外
        '    wDate2 = MXA030_翌営業年月日計算(wDate2, p支払日, p営業日区分) ' 07/01/30 V180
        'End If
        
        'wdd = Day(wDate2)
        
        'If (wdd <> p支払日 And p支払日 <> 31) _
        '   Or (wdd <> p支払日 And p判断 = "*") Then    '2004/03/28
        '    If wDate2 = p年月日 Then
        '        MXA030_実行支払年月 = DateAdd("m", -1, MXA030_実行支払年月)
        '    End If
        'End If
    End If
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MXA030_実行支払年月_ERR:
    pERR_MES = pPROGRAM_ID + "/ MXA030_実行支払年月() でエラー" + vbCrLf + vbCrLf + _
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
' MXA030_金利初回年月
'------------------------------------------------
Public Function MXA030_金利初回年月(p利息区分 As String, p利息支払 As Integer, p支払日 As Integer, p営業日 As Integer, p実行日 As Variant, p初回返済年月 As Variant, p返済単位 As Integer) As Variant
'
    Dim wi01 As Integer
    Dim wd01 As Date
    Dim wv01 As Variant, wv02 As Variant
'
    On Error GoTo MXA030_金利初回年月_ERR
'
    If p利息区分 = XMXA020_区分("利息区分", "利息先払") Then
    '実行日の年月
        
        If IsNull(p実行日) Then
            MXA030_金利初回年月 = Null
            Exit Function
        End If
        
        If CStr(p利息支払) = XMXA020_区分("利息支払", "毎月") Then
            MXA030_金利初回年月 = Format(CDate(p実行日), "yyyy/mm/dd")
        Else
            wv01 = Format(p初回返済年月, "yyyy/mm/01")
            Do While Format(CDate(wv01), "yyyy/mm/01") >= Format(CDate(p実行日), "yyyy/mm/01")
                wv02 = DateAdd("m", -p返済単位, CDate(wv01))
            
                If Format(CDate(wv02), "yyyy/mm/01") >= Format(CDate(p実行日), "yyyy/mm/01") Then
                    wv01 = wv02
                Else
                    Exit Do
                End If
            Loop
            
            MXA030_金利初回年月 = Format(CDate(wv01), "yyyy/mm/dd")
        
        End If
    
    ElseIf p利息区分 = XMXA020_区分("利息区分", "利息後払") Then
        
        If IsNull(p初回返済年月) Then
            MXA030_金利初回年月 = Null
            Exit Function
        End If
            
        If CStr(p利息支払) = XMXA020_区分("利息支払", "毎月") Then
        
            If IsNull(p実行日) Then
                MXA030_金利初回年月 = Null
                Exit Function
            End If
            
            '2010/01/10 変更
'            '実行支払年月
'            wv01 = MXA030_実行支払年月(p実行日, p支払日, p営業日, "=")
'            '据置回数
'            wi01 = DateDiff("m", P8.FCDate(wv01), p初回返済年月)
'
'            If wi01 < 2 Then
'            '据置回数<2  初回返済年月
'                MXA030_金利初回年月 = Format(CDate(p初回返済年月), "yyyy/mm/dd")
'
'            Else
'            '据置回数>=2 2回目の返済年月
'                'プラス1月で2回目の返済年月
'                wd01 = DateAdd("m", 1, CDate(wv01))
'                MXA030_金利初回年月 = wd01
'
'            End If
        
            MXA030_金利初回年月 = Format(CDate(p実行日), "yyyy/mm/dd")
        
        ElseIf CStr(p利息支払) = XMXA020_区分("利息支払", "一括") Then
        '初回返済年月
            
            '2010/01/10 変更
            'MXA030_金利初回年月 = Format(CDate(p初回返済年月), "yyyy/mm/dd")
            
            wv01 = Format(p初回返済年月, "yyyy/mm/01")
            Do While Format(CDate(wv01), "yyyy/mm/01") >= Format(CDate(p実行日), "yyyy/mm/01")
                wv02 = DateAdd("m", -p返済単位, CDate(wv01))
            
                If Format(CDate(wv02), "yyyy/mm/01") >= Format(CDate(p実行日), "yyyy/mm/01") Then
                    wv01 = wv02
                Else
                    Exit Do
                End If
            Loop
            
            MXA030_金利初回年月 = Format(CDate(wv01), "yyyy/mm/dd")
            
        End If
    End If
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MXA030_金利初回年月_ERR:
    pERR_MES = pPROGRAM_ID + "/ MXA030_金利初回年月() でエラー" + vbCrLf + vbCrLf + _
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
' MXA030_ReportColor
'------------------------------------------------
Public Function MXA030_ReportColor(pControl As Field, Optional pColor As String = "")
'
    If P8.FCDbl(pControl) < 0 Then
        pControl.ForeColor = C_Red
    Else
        If pColor = "black" Then
            pControl.ForeColor = C_Black
        ElseIf pColor = "blue" Then
            pControl.ForeColor = C_Blue
        ElseIf pColor = "green" Then
            pControl.ForeColor = C_Green
        Else
            pControl.ForeColor = C_Black
        End If
    End If
'
End Function

'------------------------------------------------
' MXA030_100Percent
'------------------------------------------------
Public Function MXA030_100Percent(pDouble As Double, pFormatType As String) As String
'
    On Error GoTo MXA030_100Percent_ERR
'
    If pDouble = 100 Then
        MXA030_100Percent = ""
    ElseIf pDouble = 0 Then
        MXA030_100Percent = "0%"
    Else
        MXA030_100Percent = P8.FFormat(pDouble / 100, pFormatType)
    End If
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MXA030_100Percent_ERR:
    pERR_MES = pPROGRAM_ID + "/ MXA030_100Percent() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function
'
'/**********************************************************************************
'/ Modyule Name    : ファイル削除
'/ Argument        : IN      pObjName   As String
'/                 : OUT     INTEGER    TRUE / FALSE
'/**********************************************************************************
Public Function MXA030_Delete(pObjName As String, Optional pDrive As String = "") As Integer
'
    Dim stName As String
    Dim ws01 As String
'
    MXA030_Delete = False
'
    If pDrive = "" Then
        stName = GSerDir & "\" & pObjName
    Else
        stName = pDrive & "\" & pObjName
    End If
'
    ws01 = Dir(stName)
    If ws01 = "" Then
        Exit Function
    End If
'
    ' エラー75 回避の為のDir関数を別のパスで取得
    ws01 = Dir(App.Path)
    
    Kill stName
'
    MXA030_Delete = True
'
End Function
'
'/**********************************************************************************
'/ Modyule Name    : ファイル名Rename
'/ Modyule ID      : MXA030_Rename
'/ Argument        : IN      pObjName   As String
'/                 :         pNewName   As String
'/                 : OUT     INTEGER    TRUE / FALSE
'/**********************************************************************************
Public Function MXA030_Rename(pObjName As String, pNewName As String) As Integer
'
    Dim stSrcName As String, stDestName As String
    Dim ws01 As String
'
    MXA030_Rename = False
'
    stSrcName = GSerDir & "\" & pObjName
    stDestName = GSerDir & "\" & pNewName
'
    ws01 = Dir(stSrcName)
    If ws01 = "" Then
        Exit Function
    End If
'
    ws01 = Dir(stDestName)
    If ws01 <> "" Then
        ' エラー75 回避の為のDir関数を別のパスで取得
        ws01 = Dir(App.Path)
        
        Kill stDestName
    End If
'
    Name stSrcName As stDestName
'
    MXA030_Rename = True
'
End Function
'
'/**********************************************************************************
'/ Modyule Name    : DB最適化
'/ Argument        : IN      pObjName   As String
'/                 : OUT     Integer    TRUE / FALSE
'/**********************************************************************************
Public Function MXA030_CompactDb(pObjName As String) As Boolean
'
    On Error GoTo MXA030_CompactDb_ERR
'
    Dim wJET As New JetEngine
    Dim wsNewObjname As String, ws01 As String
    Dim stSrcName As String, stDestName As String
'
    MXA030_CompactDb = False
'
    stSrcName = GSerDir + "\" + pObjName

    ws01 = Dir(stSrcName)
    If ws01 = "" Then
        Exit Function
    End If
'
    wsNewObjname = "WK最適化" & Format(Now, "YYMMDD") & pObjName
    stDestName = GSerDir + "\" + wsNewObjname
    
    ws01 = Dir(stDestName)
    If ws01 <> "" Then
        Exit Function
    End If
    
        wJET.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0" & _
                                ";Data Source=" & stSrcName & _
                                ";Persist Security Info=False" & _
                                ";Jet OLEDB:Database Password=" & GPwd, _
                            "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                ";Data Source=" & stDestName & _
                                ";Jet OLEDB:Database Password=" & GPwd

    On Error GoTo 0
'
    GRet = MXA030_Delete(pObjName)
    If GRet = True Then
        GRet = MXA030_Rename(wsNewObjname, pObjName)
    End If
'
    MXA030_CompactDb = True
'
Exit Function
'
MXA030_CompactDb_ERR:
    Exit Function
End Function
'
'/**********************************************************************************
'/ Modyule Name    : ログファイル出力
'/ Modyule ID      : PUT_LOG_FILE
'/ Date            : 2001/01/01
'/ Written         : S.Henmi
'/ Argument        : IN      LOG MESSAGE AS STRING
'/                 : OUT     INTEGER    TRUE / FALSE
'/**********************************************************************************
Public Function PUT_LOG_FILE(LOG_MES As String) As Integer
'
    Dim LP As Integer, REC_NO As Long, FREE_FILE As Variant, FILE_CAP As Long
    Dim ws01 As String, EDT_LOG As String
    Dim pPROGRAM_ID As String
'
    On Error GoTo PUT_LOG_FILE_ERR
'
    PUT_LOG_FILE = False
    pPROGRAM_ID = "PUT_LOG_FILE"
'
'----------< REPLACE SPACE FROM &H0A,&H0D >-----------------------------------------
    EDT_LOG = ""
    For LP = 1 To Len(LOG_MES)
        If Mid(LOG_MES, LP, 1) = Chr(&HD) Then
           EDT_LOG = EDT_LOG + " "
          Else
           If Mid(LOG_MES, LP, 1) = Chr(&HA) Then
             Else
              EDT_LOG = EDT_LOG + Mid(LOG_MES, LP, 1)
           End If
        End If
    Next LP
'
'----------< FILE GET >-------------------------------------------------------------
    FREE_FILE = FreeFile
    Open GSerDir & "\" + pSYSLOG_NAME For Random As #FREE_FILE Len = Len(pREC_SYSLOG)
    FILE_CAP = LOF(FREE_FILE)
    Close #FREE_FILE
'
'----------< LOG FILE CREATE >------------------------------------------------------
    If FILE_CAP = 0 Then
       Open GSerDir & "\" + pSYSLOG_NAME For Random As #FREE_FILE Len = Len(pREC_SYSLOG)
       For LP = 1 To pREC_LOGCAP
           pREC_SYSLOG.LOG_DATE = Space(10)
           pREC_SYSLOG.FILLER00 = Space(1)
           pREC_SYSLOG.LOG_TIME = Space(8)
           pREC_SYSLOG.FILLER01 = Space(1)
           pREC_SYSLOG.LOG_MESS = Space(106)
           pREC_SYSLOG.LOG_CR = &HD
           pREC_SYSLOG.LOG_LF = &HA
           Put #FREE_FILE, LP, pREC_SYSLOG
       Next LP
       pREC_SYSLOG.LOG_DATE = "0000000000"
       Put #FREE_FILE, pREC_LOGCAP + 1, pREC_SYSLOG
       Close #FREE_FILE
    End If
'
'---------< LOG FILE PUT >----------------------------------------------------------
    FREE_FILE = FreeFile
    Open GSerDir & "\" + pSYSLOG_NAME For Random As #FREE_FILE Len = Len(pREC_SYSLOG)
'
    Get #FREE_FILE, pREC_LOGCAP + 1, pREC_SYSLOG
    REC_NO = CLng(pREC_SYSLOG.LOG_DATE)
    REC_NO = REC_NO + 1
    If REC_NO > pREC_LOGCAP Then
       REC_NO = 1
    End If
'
    pREC_SYSLOG.LOG_DATE = Format(Date, "yyyy/mm/dd")
    pREC_SYSLOG.FILLER00 = Space(1)
    pREC_SYSLOG.LOG_TIME = Format(Time, "hh:mm:ss")
    pREC_SYSLOG.FILLER01 = "," 'Space(1)
    pREC_SYSLOG.LOG_MESS = EDT_LOG
    pREC_SYSLOG.LOG_CR = &HD
    pREC_SYSLOG.LOG_LF = &HA
    Put #FREE_FILE, REC_NO, pREC_SYSLOG
'
    ws01 = "0000000000" + CStr(REC_NO)
    ws01 = Right(ws01, 10)
    pREC_SYSLOG.LOG_DATE = ws01
    pREC_SYSLOG.FILLER00 = Space(1)
    pREC_SYSLOG.LOG_TIME = Space(8)
    pREC_SYSLOG.FILLER01 = Space(1)
    pREC_SYSLOG.LOG_MESS = Space(106)
    pREC_SYSLOG.LOG_CR = &HD
    pREC_SYSLOG.LOG_LF = &HA
    Put #FREE_FILE, pREC_LOGCAP + 1, pREC_SYSLOG
'
    Close #FREE_FILE
'
    PUT_LOG_FILE = True
'
    Exit Function
'
'----------< ERROR ROUTINE >--------------------------------------------------------
PUT_LOG_FILE_ERR:
    pERR_MES = pPROGRAM_ID + "/ PUT_LOG_FILE() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)
    
    Resume PUT_LOG_FILE_ERR_END
PUT_LOG_FILE_ERR_END:
    Exit Function
'
End Function

'------------------------------------------------
' MXA030_GET_ListFL
'------------------------------------------------
Public Sub MXA030_GET_ListFL(pdata As String)
'
    Dim M As Integer
    Dim ws01 As String, ws02 As String
'
    GStr = "": GStr_2 = ""
    
    '1st
    ws01 = "": ws02 = ""
    For M = 1 To Len(pdata)
        ws01 = Mid$(pdata, M, 1)
        If ws01 = ":" Then
            Exit For
        Else
            ws02 = ws02 & ws01
        End If
    Next
    GStr = ws02
'
    'Last
    ws01 = "": ws02 = ""
    For M = Len(pdata) To 1 Step -1
        ws01 = Mid$(pdata, M, 1)
        If ws01 = ":" Then
            Exit For
        Else
            ws02 = ws01 & ws02
        End If
    Next M
    GStr_2 = ws02
'
End Sub

'------------------------------------------------
' SET_LISTCOMBO
'------------------------------------------------
Public Function SET_LISTCOMBO(pCombo As Object, pKubun As String, pValue As String) As Integer
'
    Dim j As Integer
    Dim ws01 As String
'
    SET_LISTCOMBO = -1
'
    On Error Resume Next
'
    ws01 = XMXA020_区分(pKubun, pValue)
    If ws01 <> "" Then
        For j = 0 To pCombo.ListCount
            If pCombo.List(j) = ws01 Then
                SET_LISTCOMBO = j
                
                Exit For
            End If
        Next j
    End If
'
    On Error GoTo 0
'
End Function

'------------------------------------------------
' MXA030_null置換
'------------------------------------------------
Public Sub MXA030_null置換()
'
    Dim wstr As String
'
    wstr = "Update DBDA010_借入金"
    wstr = wstr & " SET 借入計画番号=''"
    wstr = wstr & " WHERE 借入計画番号 Is Null"
    GDb.Execute wstr

    DoEvents
'
    wstr = "Update DBDA010_貸付金"
    wstr = wstr & " SET 借入計画番号=''"
    wstr = wstr & " WHERE 借入計画番号 Is Null"
    GDb.Execute wstr

    DoEvents
'
    wstr = "Update DBDA010_分岐点借入金"
    wstr = wstr & " SET 借入計画番号=''"
    wstr = wstr & " WHERE 借入計画番号 Is Null"
    GDb.Execute wstr

    DoEvents
'
    wstr = "Update DBDA010_分岐点貸付金"
    wstr = wstr & " SET 借入計画番号=''"
    wstr = wstr & " WHERE 借入計画番号 Is Null"
    GDb.Execute wstr

    DoEvents
'
    wstr = "Update DBDA010_借入金"
    wstr = wstr & " SET 金融リストラ番号=''"
    wstr = wstr & " WHERE 金融リストラ番号 Is Null"
    GDb.Execute wstr

    DoEvents
'
    wstr = "Update DBDA010_貸付金"
    wstr = wstr & " SET 金融リストラ番号=''"
    wstr = wstr & " WHERE 金融リストラ番号 Is Null"
    GDb.Execute wstr

    DoEvents
'
    wstr = "Update DBDA010_分岐点借入金"
    wstr = wstr & " SET 金融リストラ番号=''"
    wstr = wstr & " WHERE 金融リストラ番号 Is Null"
    GDb.Execute wstr

    DoEvents
'
    wstr = "Update DBDA010_分岐点貸付金"
    wstr = wstr & " SET 金融リストラ番号=''"
    wstr = wstr & " WHERE 金融リストラ番号 Is Null"
    GDb.Execute wstr

    DoEvents
'
End Sub

'/**********************************************************************************
'/ Modyule Nname   : システムログファイル出力
'/ Modyule ID      : PUT_LOG
'/ Date            : 2001/01/01
'/ Written         : S.Henmi
'/ Argument        : IN      LOG MESSAGE AS STRING
'/                 : OUT     INTEGER    TRUE / FALSE
'/**********************************************************************************
Public Function PUT_LOG(LOG_MES As String) As Integer
'
    Dim LP As Integer, REC_NO As Long, FREE_FILE As Variant, FILE_CAP As Long
    Dim ws01 As String, EDT_LOG As String
    Dim FLG_FIR As Boolean
    Dim wi01 As Integer
'
    On Error GoTo PUT_LOG_ERR
'
    PUT_LOG = False
'
'----------< FILE GET >-------------------------------------------------------------
    FREE_FILE = FreeFile
    Open GSerDir + pERRLOG_NAME For Random As #FREE_FILE Len = Len(pREC_SYSLOG)
    FILE_CAP = LOF(FREE_FILE)
    Close #FREE_FILE
'
'----------< LOG FILE CREATE >------------------------------------------------------
    If FILE_CAP = 0 Then
       Open GSerDir + pERRLOG_NAME For Random As #FREE_FILE Len = Len(pREC_SYSLOG)
       For LP = 1 To pREC_LOGCAP
           pREC_SYSLOG.LOG_DATE = Space(10)
           pREC_SYSLOG.FILLER00 = Space(1)
           pREC_SYSLOG.LOG_TIME = Space(8)
           pREC_SYSLOG.FILLER01 = Space(1)
           pREC_SYSLOG.LOG_MESS = Space(106)
           pREC_SYSLOG.LOG_CR = &HD
           pREC_SYSLOG.LOG_LF = &HA
           Put #FREE_FILE, LP, pREC_SYSLOG
       Next LP
       pREC_SYSLOG.LOG_DATE = "0000000000"
       Put #FREE_FILE, pREC_LOGCAP + 1, pREC_SYSLOG
       Close #FREE_FILE
    End If
'
'---------< LOG FILE PUT >----------------------------------------------------------
    FREE_FILE = FreeFile
    Open GSerDir + pERRLOG_NAME For Random As #FREE_FILE Len = Len(pREC_SYSLOG)
'
    Get #FREE_FILE, pREC_LOGCAP + 1, pREC_SYSLOG
    REC_NO = CLng(pREC_SYSLOG.LOG_DATE)
    REC_NO = REC_NO + 1
    If REC_NO > pREC_LOGCAP Then
       REC_NO = 1
    End If
'
    FLG_FIR = False
    '----------< REPLACE SPACE FROM &H0A,&H0D >-------------------------------------
    EDT_LOG = ""
    For LP = 1 To Len(LOG_MES)
        If Mid(LOG_MES, LP, 1) = Chr(&HD) Then
        
           If EDT_LOG <> "" Then
              If FLG_FIR = False Then
    
                 pREC_SYSLOG.LOG_DATE = Format(Date, "yyyy/mm/dd")
                 pREC_SYSLOG.FILLER00 = Space(1)
                 pREC_SYSLOG.LOG_TIME = Format(Time, "hh:mm:ss")
                 pREC_SYSLOG.FILLER01 = Space(1)
                 pREC_SYSLOG.LOG_MESS = EDT_LOG
                 pREC_SYSLOG.LOG_CR = &HD
                 pREC_SYSLOG.LOG_LF = &HA
                 Put #FREE_FILE, REC_NO, pREC_SYSLOG
                
                 REC_NO = REC_NO + 1
                 
                 FLG_FIR = True
    
              Else
                 ws01 = EDT_LOG
                 wi01 = 1
                 
                 Do While ws01 <> ""
                    ws01 = MidB(EDT_LOG, wi01, 106)
                    
                    pREC_SYSLOG.LOG_DATE = Space(10)
                    pREC_SYSLOG.FILLER00 = Space(1)
                    pREC_SYSLOG.LOG_TIME = Space(8)
                    pREC_SYSLOG.FILLER01 = Space(1)
                    pREC_SYSLOG.LOG_MESS = ws01
                    pREC_SYSLOG.LOG_CR = &HD
                    pREC_SYSLOG.LOG_LF = &HA
                    Put #FREE_FILE, REC_NO, pREC_SYSLOG
                    
                    REC_NO = REC_NO + 1
                 
                    wi01 = wi01 + 106
                    If wi01 >= LenB(EDT_LOG) Then
                       ws01 = ""
                    End If
                    
                 Loop
           
              End If
              
           End If
           
           EDT_LOG = ""
           
          Else
           If Mid(LOG_MES, LP, 1) = Chr(&HA) Then
              EDT_LOG = ""
              
             Else
             
              EDT_LOG = EDT_LOG + Mid(LOG_MES, LP, 1)
              
           End If
        End If
    Next LP
    '
    ws01 = "0000000000" + CStr(REC_NO)
    ws01 = Right(ws01, 10)
    pREC_SYSLOG.LOG_DATE = ws01
    pREC_SYSLOG.FILLER00 = Space(1)
    pREC_SYSLOG.LOG_TIME = Space(8)
    pREC_SYSLOG.FILLER01 = Space(1)
    pREC_SYSLOG.LOG_MESS = Space(106)
    pREC_SYSLOG.LOG_CR = &HD
    pREC_SYSLOG.LOG_LF = &HA
    Put #FREE_FILE, pREC_LOGCAP + 1, pREC_SYSLOG
'
    Close #FREE_FILE
'
    PUT_LOG = True
'
    Exit Function
'
'----------< ERROR ROUTINE >--------------------------------------------------------
PUT_LOG_ERR:
    pERR_MES = pPROGRAM_ID + "/ PUT_LOG() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)
    
    Resume PUT_LOG_ERR_END
PUT_LOG_ERR_END:
    Exit Function
'
End Function

'------------------------------------------------
' MXA030_LOG_WRITE
'------------------------------------------------
Public Sub MXA030_LOG_WRITE(pPID As String, pKubun As String, pLOG_MES As String)
'
    Dim wi01 As Integer
    Dim wstr As String

    Dim wJET As New JetEngine
    Dim wDb As New ADODB.Connection
'
    On Error GoTo MXA030_LOG_WRITE_ERR
'
    'ログ区分
    wi01 = 0
    Select Case pKubun
    Case "ログイン": wi01 = 0
    Case "追加": wi01 = 1
    Case "更新": wi01 = 2
    Case "削除": wi01 = 3
    Case "帳票": wi01 = 4
    Case "ログアウト": wi01 = 5
    Case "照会": wi01 = 6
    End Select
'
    '----------< LOG.mdb Open >------------------------------------------------
    Call AdoDbOpen("Jet", wDb, GCurDir + "\LOG.mdb", "", , GPwd)

        '----------< T_Log 更新 >----------------------------------------------
        wstr = ""
        wstr = wstr & "Insert into T_Log"
        wstr = wstr & " (ログ日付,ログ時刻,USERID,PROGRAMID,ログ区分,内容,内容2,内容3)"
        wstr = wstr & "  Values("
        wstr = wstr & "#" & Format(Date, "yyyy/mm/dd") & "#,"
        wstr = wstr & "#" & Format(Now, "hh:mm:ss") & "#,"
        wstr = wstr & "'" & GUserID & "',"
        wstr = wstr & "'" & pPID & "',"
        wstr = wstr & wi01 & ","
        wstr = wstr & "'" & Left(pLOG_MES, 240) & "',"
        wstr = wstr & "'" & Mid(pLOG_MES, 240, 240) & "',"
        wstr = wstr & "'" & Mid(pLOG_MES, 480, 240) & "'"
        wstr = wstr & ")"
        wDb.Execute (wstr)
        
    '----------< LOG.mdb Close >-----------------------------------------------
    wDb.Close
    Set wDb = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MXA030_LOG_WRITE_ERR:
    pERR_MES = pPROGRAM_ID + "/ MXA030_LOG_WRITE() でエラー" + vbCrLf + vbCrLf + _
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
' UNLOAD_REPFRM
'　帳票画面、帳票指示画面を全てアンロードする   2011.11.17 By m.mino
'------------------------------------------------
Public Sub UNLOAD_REPFRM()

    Unload frm_R帳票出力設定 '最初にUNLOAD
    
    Unload frm_R借入金台帳
    Unload frm_R借入金明細表
    Unload frm_R借入一覧表
    Unload frm_R返済予定表
    Unload frm_R借入残高表
    Unload frm_R借入残高推移表
    Unload frm_R利息前払未払明細表
    Unload frm_R利息前払未払残高表
    Unload frm_R利息残高推移表
    Unload frm_R平均金利平均残高推移表
    Unload frm_R平均金利平均残高表
    Unload frm_R金融機関別残高表
    Unload frm_R簡易資金繰表
    Unload frm_R年度別比較表
    Unload frm_R仕訳データ作成
    Unload frm_R決算仕訳データ作成
    Unload frm_R借入金時価評価一覧表
    Unload frm_R借入金時価評価適用金利一覧
    Unload frm_R借入金時価評価明細表
    Unload frm_R損益利息一覧表
    
    '杉村倉庫仕様
'    Unload frm_R1年内返済集計表
'    Unload frm_R銀行別利息表
'    Unload frm_R支払利息推移表
    
    '神姫バス
'    Unload frm_R資金繰表
'    Unload frm_R長短振替表
End Sub

'------------------------------------------------
' UNLOAD_MASFRM
'　マスタ登録画面を全てアンロードする   2011.11.19 By m.mino
'------------------------------------------------
Public Sub UNLOAD_MASFRM()

    Unload frm_R帳票出力設定 '最初にUNLOAD
    
    Unload frm_Mユーザー設定
    Unload frm_M基準金利マスタ
    Unload frm_M金利シミュレーショングループマスタ
    Unload frm_M銀行マスタ
    Unload frm_M借入種別マスタ
    Unload frm_M勘定科目マスタ
    Unload frm_M補助科目マスタ
    Unload frm_M個別補助科目マスタ
    Unload frm_M部門マスタ
    Unload frm_M長期プライムレート

End Sub

'------------------------------------------------
' UNLOAD_ALLFRM
'　全フォームをアンロードする   2011.11.26 By m.mino
'------------------------------------------------
Public Sub UNLOAD_ALLFRM()

    Call UNLOAD_MASFRM
    Call UNLOAD_REPFRM
    
    Unload frm_Fログ照会
    Unload frm_F借入金明細表
    Unload frm_F借入登録データ照会
    Unload frm_I金利シミュレーション入力
    'Unload frm_I銀行登録
    Unload frm_I借入金登録
    Unload frm_I借入金登録_金利変更
    Unload frm_I借入金登録_銀行
    Unload frm_I借入金登録_内入
    Unload frm_I借入金登録_明細
    Unload frm_K借入金検索
    
End Sub

'------------------------------------------------
' MXA030_GRPTCLEAR
'------------------------------------------------
Public Sub MXA030_GRPTCLEAR()
'
    GRpt.推移 = ""
    GRpt.選択 = ""
    GRpt.実績 = ""
    GRpt.作業 = ""
    GRpt.集計 = ""
    GRpt.指定 = ""
    
    GRpt.連結売上 = ""
    GRpt.売上 = ""
    GRpt.借入 = ""
    GRpt.設備 = ""
    GRpt.金融 = ""
    GRpt.設備R = ""
    GRpt.リス = ""
    
    GRpt.連結売上2 = ""
    GRpt.売上2 = ""
    GRpt.借入2 = ""
    GRpt.設備2 = ""
    GRpt.金融2 = ""
    GRpt.設備R2 = ""
    GRpt.リス2 = ""
    
    GRpt.テキスト_01 = ""
    GRpt.テキスト_02 = ""
    
    GRpt.借入金管理区分 = 0
    GRpt.詳細表示 = 0
    GRpt.CSV = 0
    GRpt.千円単位 = 0
    GRpt.金利SM = 0
      
    GRpt.チェック_01 = 0
    GRpt.チェック_02 = 0
    GRpt.チェック_03 = 0
    GRpt.チェック_04 = 0

    G金利SM = False
    
    GRpt.C_種別 = ""
    GRpt.C_部門 = ""
    GRpt.C_金融 = ""
    GRpt.C_銀行 = ""

    GRpt.S_利息 = ""
    GRpt.S_種別 = ""
    GRpt.S_部門 = ""
    GRpt.S_金融 = ""
    GRpt.S_銀行 = ""
    GRpt.S_金利 = ""

    GRpt.NewPage1 = 0
    GRpt.NewPage2 = 0
    GRpt.NewPage3 = 0
    GRpt.NewPage4 = 0
'
End Sub

'------------------------------------------------
'　UNLOAD_借入金FRM
'------------------------------------------------
Public Sub UNLOAD_借入金FRM()
'
    Dim f As Form
'
    For Each f In Forms
        If (f Is frm_I借入金登録 And f.Visible = True) _
        Or (f Is frm_I借入金登録_銀行 And f.Visible = True) _
        Or (f Is frm_I借入金登録_明細 And f.Visible = True) _
        Or (f Is frm_I借入金登録_内入 And f.Visible = True) _
        Or (f Is frm_K借入金検索 And f.Visible = True) Then
            MsgBox "借入金登録のフォームを閉じます。", vbInformation
            Exit For
        End If
    Next
    
    Unload frm_I借入金登録
    Unload frm_I借入金登録_金利変更
    Unload frm_I借入金登録_銀行
    Unload frm_I借入金登録_内入
    Unload frm_I借入金登録_明細
    Unload frm_K借入金検索
'
End Sub
