Attribute VB_Name = "MAA001_Main"
Option Explicit
'
Private Const pPROGRAM_ID As String = "MAA001_Main"
'
Dim wdate As Date

Public Sub Main()
    ' =========================================
    '            ２重起動禁止
    ' =========================================
    If App.PrevInstance Then
        GRet = MsgBox("このプログラムは、すでに実行されています", vbOKOnly + vbCritical)
        End
    End If
    
    ' =========================================
    '            コンピュータ名 取得
    ' =========================================
    GLong1 = 16
    GStr = String(GLong1, Chr(32))
    GRet = GetComputerName(GStr, GLong1)
    GStr = Trim(GStr)
    GLong1 = Len(GStr)
    
    GMyComputerName = Left(GStr, GLong1 - 1)
    ' =========================================
    '            システム初期設定
    ' =========================================
    GCurDir = App.Path
    GPwd = "inkinhkheshh2IHPDKPI"
    
    '金剛石 or 借換たろう！
    'GProduct = "金剛石"
    GProduct = "借換たろう！"
    
    If GProduct = "借換たろう！" Then
    '借換たろう！ 製品選択
        GSys.Sys = "借入金"
'        GSys.Sys = "借入金 お試し版"
        'GSys.Sys = "借入金 Lite"
    End If

    GUserID = ""
    GUserKen = 9
'
    '*** EnterKeey Init ***
    Call CEkey.X010_InitSetAllSelect("ZU020_ComboBox")
    Call CEkey.X011_InitSetNextTab("ZU020_ComboBox", "ZU030_ListBox", "ZU050_Button")
'
    ' =========================================
    '  借換たろう！お試し版帳票出力回数チェック
    ' =========================================
    If GSys.Sys = "借入金 お試し版" Then
        GPriCnt = MAA001_KARIKAETAROU_CHECK()
        If GPriCnt > 100 Then
            GRet = MsgBox("借換たろう！お試し版使用回数が過ぎました。" + Chr(13) + vbCrLf + "借換たろう！を終了します。", vbOKOnly + vbExclamation)
            
            End
        End If
    End If
'
    ' =========================================
    '           CHECK SERIAL.FIL 有効日付
    ' =========================================
    If GSys.Sys = "借入金 お試し版" Then
        GDate1 = CDate("2010/12/31")
        GRet = MAA100_KARIKAETAROU()
        If GRet <> True Then
            GRet = MsgBox("借換たろう！を終了します。", vbOKOnly + vbCritical)
            End
        End If
        wdate = GDate1
        
    ElseIf GSys.Sys = "借入金 Lite" Then
    
    'ElseIf GSys.Sys = "借入金" Then
        'GRet = MAA100_SERIAL()
        'If GRet <> True Then
        '    GRet = MsgBox("シリアル情報が正しくありません。" + Chr(13) + vbCrLf + GProduct + "を終了します", vbOKOnly + vbCritical)
        '    End
        'End If
    
    'シリアル無し
    ElseIf GSys.Sys = "借入金" Then
        GDate1 = CDate("2010/12/31")
        wdate = GDate1
    
    Else
    '金剛石
        GRet = MAA100_SERIAL()
        If GRet <> True Then
            GRet = MsgBox("シリアル情報が正しくありません。" + Chr(13) + vbCrLf + GProduct + "を終了します", vbOKOnly + vbCritical)
            End
        End If
    End If
'
    ' =========================================
    '           SET SERIAL情報
    ' =========================================
    If GSys.Sys = "借入金 お試し版" Then
    '借入金 お試し版
        GSys.Sys = "借入金 お試し版"
        GSys.Mem = "単一"
        GSys.Sit = False
        GSys.Lan = False
        GSys.Ker = False
        GSys.Han = False
    
    ElseIf GSys.Sys = "借入金 Lite" Then
    '借入金 Lite
        GSys.Sys = "借入金 Lite"
        GSys.Mem = "単一"
        GSys.Sit = False
        GSys.Lan = False
        GSys.Ker = False
        GSys.Han = False
    
    ElseIf GSys.Sys = "借入金" Then
    '借入金
        GSys.Mem = "複数"
        GSys.Sit = True
        GSys.Lan = True
        GSys.Ker = False
        GSys.Han = False
    Else
        Call MXA010_GSys設定
    End If
    
'
    ' =========================================
    '           CHECK SET SerVer
    ' =========================================
    GRet = MAA100_MDBVER() '借入金 お試し版選択あり
    If GRet <> True Then
        GRet = MsgBox("バージョンが正しくありません。" + Chr(13) + vbCrLf + GProduct + "を終了します", vbOKOnly + vbCritical)
        End
    End If
'

    ' =========================================
    '           Csv File Drive
    ' =========================================
    Call MX040_CsvPath
'
    ' =========================================
    '           Main Form Show
    ' =========================================
    'frm_Fメインフォーム.Show
'    FAA020_ログインユーザ選択.Show
    frm_Fデータベース選択.Show
'
End Sub

'------------------------------------------------
' MAA001_KARIKAETAROU_CHECK
'------------------------------------------------
Private Function MAA001_KARIKAETAROU_CHECK() As Integer
'
    Dim wJET As New JetEngine
    Dim wDb As New ADODB.Connection
    Dim wRs3 As ADODB.Recordset
    Dim wstr As String
'
    On Error GoTo MAA001_KARIKAETAROU_CHECK_ERR
'
    MAA001_KARIKAETAROU_CHECK = 800
'
    Call AdoDbOpen("Jet", wDb, GCurDir + "\" + GMain, "", , GPwd)

    wstr = "Select * From 借換たろうシステム制御"
    wstr = wstr + " Where System = 'System'"
    Call AdoRecordsetOpen(wDb, wRs3, wstr)
    If Not wRs3.eof Then
        MAA001_KARIKAETAROU_CHECK = wRs3("出力回数")
    End If
    wRs3.Close
    Set wRs3 = Nothing
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA001_KARIKAETAROU_CHECK_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA001_KARIKAETAROU_CHECK() でエラー" + vbCrLf + vbCrLf + _
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
' MAA001_KARIKAETAROU_CNT
'------------------------------------------------
Public Sub MAA001_KARIKAETAROU_CNT()
'
    Dim wJET As New JetEngine
    Dim wDb As New ADODB.Connection
    Dim wRs3 As ADODB.Recordset
    
    Dim wstr As String
'
    On Error GoTo MAA001_KARIKAETAROU_CNT_ERR
'
    Call AdoDbOpen("Jet", wDb, GCurDir + "\" + GMain, "", , GPwd)

    wstr = "Select * From 借換たろうシステム制御"
    wstr = wstr + " Where System = 'System'"
    Call AdoRecordsetOpen(wDb, wRs3, wstr)
    If Not wRs3.eof Then
        GPriCnt = GPriCnt + 1
        wRs3("出力回数") = GPriCnt
    
        wRs3.Update
    End If
    wRs3.Close
    Set wRs3 = Nothing
'
    Call MAA001_KARIKAETAROU_PRE
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA001_KARIKAETAROU_CNT_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA001_KARIKAETAROU_CNT() でエラー" + vbCrLf + vbCrLf + _
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
' MAA001_KARIKAETAROU_PRE
'------------------------------------------------
Public Sub MAA001_KARIKAETAROU_PRE()
'
    Dim wi01 As Integer
'
    wi01 = 100 - GPriCnt
    If GPriCnt > 100 Then
        frm_Fメインフォーム.L_回数.Caption = ""
    Else
        frm_Fメインフォーム.L_回数.Caption = "帳票出力 残り" & wi01 & "回"
    End If
'
End Sub

