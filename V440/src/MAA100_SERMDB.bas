Attribute VB_Name = "MAA100_SERMDB"
Option Explicit
'
Private Const pPROGRAM_ID As String = "MAA100_SERMDB"

'------------------------------------------------
' MAA100_SERIAL
'------------------------------------------------
Public Function MAA100_SERIAL() As Boolean
'
    Dim P1 As String, P2 As String
    Dim ws01 As String
    Dim wdate As Date
'
    On Error GoTo MAA100_SERIAL_ERR
'
    MAA100_SERIAL = False
'
    '借換たろう！
    If GSys.Sys = "借入金 お試し版" Or GSys.Sys = "借入金 Lite" Then
    '借入金 お試し版
    '借入金 Lite
        MAA100_SERIAL = True
        Exit Function
    End If
'
    ' =========================================
    '           SERIAL.FILのチェック
    ' =========================================
    GRet = SERIAL_READ_MAIN(P1, P2)
    If GRet <> True Then
        Exit Function
    End If
'
    ' =========================================
    '           有効日日付のＣＨＥＣＫ
    ' =========================================
    ws01 = Mid$(P2, 2, 6)
    If ws01 <> "999999" Then
        ws01 = Left$(ws01, 2) & "/" & Mid$(ws01, 3, 2) & "/" & Right$(ws01, 2)
        wdate = CDate(ws01)
        If Format(wdate, "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
            GRet = MsgBox("有効日付が過ぎました", vbOKOnly + vbCritical)
            Exit Function
        End If
    End If
'
    GP2 = P2
'
    MAA100_SERIAL = True
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA100_SERIAL_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA100_SERIAL() でエラー" + vbCrLf + vbCrLf + _
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
' MAA100_MDBVER
'------------------------------------------------
Public Function MAA100_MDBVER(Optional pFun As String = "") As Boolean
'
    Dim wJET As New JetEngine
    Dim wDb As New ADODB.Connection
    Dim wRs3 As ADODB.Recordset

    Dim L As Integer
    Dim wl01 As Long
    
    Dim wVer_SMain As String, wVer_STemp As String
    Dim wVer_CMain As String, wVer_CTemp As String
    Dim wstr As String, wdate As String
    Dim ws01 As String, ws02 As String, ws_Msg As String
'
    On Error GoTo MAA100_MDBVER_ERR
'
    MAA100_MDBVER = False
'
    ws01 = "": ws02 = ""
    GSerComputerName = GMyComputerName
    GSerDir = GCurDir
    ws_Msg = "バージョンの違うMDBを参照しています。"
'
    '----------< Set GSerDir GSerComputerName >-------------------------------------
    '----------< GCurDir GTemp Open >-----------------------------------------------
    Call AdoDbOpen("Jet", wDb, GCurDir + "\" + GTemp, "", , GPwd)

    wstr = "Select * From DAAA020_コントロール"
    wstr = wstr + " Where System = 'System'"
    Call AdoRecordsetOpen(wDb, wRs3, wstr)
    If Not wRs3.EOF Then
        ws01 = P8.FCStr(wRs3("サーバーフォルダ"))
        
        If GSys.Lan = False Then
            If ws01 = "" Or ws01 <> GCurDir Then
                wRs3("サーバーフォルダ") = GCurDir
                GSerDir = GCurDir
            
                wRs3.Update
            End If
        Else
            If ws01 <> "" Then
                GSerDir = ws01
            End If
        End If
    End If
    wRs3.Close
    Set wRs3 = Nothing
'
    '----------< Check GSerDir >----------------------------------------------------
    GRet = CHECK_SERDIR
    If GRet <> True Then
        If pFun = "指定" Then
            wstr = ""
            wstr = "UPDATE DAAA020_コントロール"
            wstr = wstr + " SET サーバーフォルダ ='" & GCurDir & "'"
            wstr = wstr + " Where System = 'System'"
            wDb.Execute wstr
        Else
        
            GRet = MsgBox("サーバーに接続できません。ローカルで起動します。", vbOKOnly + vbInformation)
        End If
        
        GSerDir = GCurDir
        GSerComputerName = GMyComputerName
    End If
    
    '----------< GCurDir GTemp Close >----------------------------------------------
    wDb.Close
    Set wDb = Nothing
'
    If GSys.Sys = "借入金 お試し版" Then
    '借入金 お試し版

        '借換たろう！DB情報
        GDbName = GSerDir + "\" + GTemp
    
    Else
        If GSys.Mem = "単一" Then
            For L = 1 To 1000
                GDbName = GSerDir + "\K" + Format(L, "000") + ".mdb"
                If Dir(GDbName) <> "" Then
                    Exit For
                End If
            Next
        End If
    
        If L >= 1000 And GSys.Mem = "単一" Then
            GDbName = GSerDir + "\K001.Mdb"
            wdate = Format(Date, "yyyy/mm/dd")
    
            wJET.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0" & _
                                    ";Data Source=" & GSerDir & "\" & GTemp & _
                                    ";Persist Security Info=False" & _
                                    ";Jet OLEDB:Database Password=" & GPwd, _
                                    "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                    ";Data Source=" & GSerDir & "\K001.mdb" & _
                                    ";Jet OLEDB:Database Password=" & GPwd
    
            '----------< GSerDir K001.mdb Open >----------------------------------------
            Call AdoDbOpen("Jet", wDb, GSerDir + "\K001.Mdb", "", , GPwd)
    
                '----------< 企業名マスタ 更新 >----------------------------------------
                wstr = ""
                wstr = wstr + "Insert into "
                wstr = wstr + "  DAAA070_企業名マスタ"
                wstr = wstr + "  (企業名Key,企業名,最新処理日,作成日,DB名)"
                wstr = wstr + "  Values("
                wstr = wstr + "'---------------',"
                wstr = wstr + "'---------------',"
                wstr = wstr + "#" + wdate + "#,"
                wstr = wstr + "#" + wdate + "#,"
                wstr = wstr + "'K001.mdb'"
                wstr = wstr + ")"
                wDb.Execute (wstr)
    
            '----------< GSerDir K001.mdb Close >---------------------------------------
            wDb.Close
            Set wDb = Nothing
    
            '----------< GSerDir GMain Open >-------------------------------------------
            Call AdoDbOpen("Jet", wDb, GSerDir + "\" + GMain, "", , GPwd)
    
                '----------< 企業名マスタ 更新 >----------------------------------------
                wDb.Execute (wstr)
    
            '----------< GSerDir GMain Close >------------------------------------------
            wDb.Close
            Set wDb = Nothing
        End If
    
        If GSys.Mem = "複数" Then
            GDbName = GSerDir + "\" + GTemp
        End If
    End If
'

'
    ' =========================================
    '           Check Version mdb
    ' =========================================
    ws01 = "": ws02 = ""
    wVer_SMain = "": wVer_STemp = ""
    wVer_CMain = "": wVer_CTemp = ""
    
    '----------< GSerDir GMain Open >-----------------------------------------------
    Call AdoDbOpen("Jet", wDb, GSerDir + "\" + GMain, "", , GPwd)
        
    wstr = "Select * From LIST000_データ保存先マスタ"
    wstr = wstr + " Where System = 'System'"
    Call AdoRecordsetOpen(wDb, wRs3, wstr)
    If Not wRs3.EOF Then
        wVer_SMain = wRs3("Version")
    End If
    wRs3.Close
    Set wRs3 = Nothing
    
    '----------< GSerDir GMain Close >----------------------------------------------
    wDb.Close
    Set wDb = Nothing
    
    '----------< GSerDir GTemp Open >-----------------------------------------------
    Call AdoDbOpen("Jet", wDb, GSerDir + "\" + GTemp, "", , GPwd)
        
    wstr = "Select * From DAAA000_バージョン"
    wstr = wstr + " Where System = 'System'"
    Call AdoRecordsetOpen(wDb, wRs3, wstr)
    If Not wRs3.EOF Then
        wVer_STemp = wRs3("Version")
    End If
    wRs3.Close
    Set wRs3 = Nothing
    
    '----------< GSerDir GTemp Close >----------------------------------------------
    wDb.Close
    Set wDb = Nothing
'
    '----------< GSerDir GMain × GSerDir GTemp >-----------------------------------
    If wVer_SMain <> wVer_STemp Then
        ws01 = GSerComputerName & Space(1) & Left$(GMain, 1) & ":" & wVer_SMain
        ws01 = ws01 & " , " & Left$(GTemp, 1) & ":" & wVer_STemp
        GRet = MsgBox(ws_Msg & vbCrLf & ws01 & vbCrLf, vbOKOnly + vbCritical)
        Exit Function
    End If
'
    
'
    If GCurDir <> GSerDir Then
        '----------< GCurDir GMain Open >-------------------------------------------
        Call AdoDbOpen("Jet", wDb, GCurDir + "\" + GMain, "", , GPwd)
        
        wstr = "Select * From LIST000_データ保存先マスタ"
        wstr = wstr + " Where System = 'System'"
        Call AdoRecordsetOpen(wDb, wRs3, wstr)
        If Not wRs3.EOF Then
            wVer_CMain = wRs3("Version")
        End If
        wRs3.Close
        Set wRs3 = Nothing
    
        '----------< GCurDir GMain Close >------------------------------------------
        wDb.Close
        Set wDb = Nothing
    
        '----------< GCurDir GTemp Open >-------------------------------------------
        Call AdoDbOpen("Jet", wDb, GCurDir + "\" + GTemp, "", , GPwd)
        
        wstr = "Select * From DAAA000_バージョン"
        wstr = wstr + " Where System = 'System'"
        Call AdoRecordsetOpen(wDb, wRs3, wstr)
        If Not wRs3.EOF Then
            wVer_CTemp = wRs3("Version")
        End If
        wRs3.Close
        Set wRs3 = Nothing
    
        '----------< GCurDir GTemp Close >------------------------------------------
        wDb.Close
        Set wDb = Nothing
'
        '----------< GCurDir GMain × GCurDir GTemp >-------------------------------
        If wVer_CMain <> wVer_CTemp Then
            ws01 = GMyComputerName & Space(1) & Left$(GMain, 1) & ":" & wVer_CMain
            ws01 = ws01 & " , " & Left$(GTemp, 1) & ":" & wVer_CTemp
            GRet = MsgBox(ws_Msg & vbCrLf & ws01 & vbCrLf, vbOKOnly + vbCritical)
            Exit Function
        End If
        
        '----------< GSerDir GMain × GCurDir GMain >-------------------------------
        If wVer_SMain <> wVer_CMain Then
            ws01 = GSerComputerName & Space(1) & Left$(GMain, 1) & ":" & wVer_SMain
            ws02 = GMyComputerName & Space(1) & Left$(GMain, 1) & ":" & wVer_CMain
            GRet = MsgBox(ws_Msg & vbCrLf & ws01 & vbCrLf & ws02 & vbCrLf, vbOKOnly + vbCritical)
            Exit Function
        End If
        
        '----------< GSerDir GTemp × GCurDir GTemp >-------------------------------
        If wVer_STemp <> wVer_CTemp Then
            ws01 = GSerComputerName & Space(1) & Left$(GTemp, 1) & ":" & wVer_STemp
            ws02 = GMyComputerName & Space(1) & Left$(GTemp, 1) & ":" & wVer_CTemp
            GRet = MsgBox(ws_Msg & vbCrLf & ws01 & vbCrLf & ws02 & vbCrLf, vbOKOnly + vbCritical)
            Exit Function
        End If
        
        '----------< GSerDir GMain × GCurDir GTemp >-------------------------------
        If wVer_SMain <> wVer_CTemp Then
            ws01 = GSerComputerName & Space(1) & Left$(GMain, 1) & ":" & wVer_SMain
            ws02 = GMyComputerName & Space(1) & Left$(GTemp, 1) & ":" & wVer_CTemp
            GRet = MsgBox(ws_Msg & vbCrLf & ws01 & vbCrLf & ws02 & vbCrLf, vbOKOnly + vbCritical)
            Exit Function
        End If
        
        '----------< GSerDir GTemp × GCurDir GMain >-------------------------------
        If wVer_STemp <> wVer_CMain Then
            ws01 = GSerComputerName & Space(1) & Left$(GTemp, 1) & ":" & wVer_STemp
            ws02 = GMyComputerName & Space(1) & Left$(GMain, 1) & ":" & wVer_CMain
            GRet = MsgBox(ws_Msg & vbCrLf & ws01 & vbCrLf & ws02 & vbCrLf, vbOKOnly + vbCritical)
            Exit Function
        End If
    Else
        wVer_CMain = wVer_SMain
        wVer_CTemp = wVer_STemp
    End If
'

'
    '----------< GSerDir KXXX.mdb Open >--------------------------------------------
    Call AdoDbOpen("Jet", wDb, GDbName, "", , GPwd)
            
    wstr = "Select * From DAAA000_バージョン"
    wstr = wstr + " Where System = 'System'"
    Call AdoRecordsetOpen(wDb, wRs3, wstr)
    If Not wRs3.EOF Then
        GVerNo = wRs3("Version")
    End If
    wRs3.Close
    Set wRs3 = Nothing
    
    '----------< GSerDir KXXX.mdb Open × GSerDir GMain >---------------------------
    If GVerNo <> wVer_SMain Then
        ws01 = GSerComputerName & " : " & GVerNo & " : " & wVer_SMain
        GRet = MsgBox(ws_Msg & vbCrLf & ws01 & vbCrLf, vbOKOnly + vbCritical)
            
        '----------< GSerDir KXXX.mdb Close >---------------------------------------
        wDb.Close
        Set wDb = Nothing
        
        Exit Function
    End If
    
    '----------< GSerDir KXXX.mdb Open × GSerDir GTemp >---------------------------
    If GVerNo <> wVer_STemp Then
        ws01 = GSerComputerName & " : " & GVerNo & " : " & wVer_STemp
        GRet = MsgBox(ws_Msg & vbCrLf & ws01 & vbCrLf, vbOKOnly + vbCritical)
            
        '----------< GSerDir KXXX.mdb Close >---------------------------------------
        wDb.Close
        Set wDb = Nothing
        
        Exit Function
    End If
    '
    ' =========================================
    '            システム日付の改ざんＣＨＥＣＫ
    ' =========================================
    ws01 = Mid$(GP2, 2, 6)
    If ws01 <> "999999" Then
        wstr = ""
        wstr = wstr + "Select * From DAAA000_バージョン"
        wstr = wstr + " Where System = 'System'"
        Call AdoRecordsetOpen(wDb, wRs3, wstr)
        If Not wRs3.EOF Then
            If P8.FCIsDate(wRs3("実行日")) Then
                If CDate(wRs3("実行日")) > Date Then
                    MsgBox "日付に不正があります"
                        
                    wRs3.Close
                    Set wRs3 = Nothing
                    '----------< GSerDir KXXX.mdb Close >-------------------------------
                    wDb.Close
                    Set wDb = Nothing
                        
                    Exit Function
                Else
                    wRs3("実行日") = Format(Date, "yyyy/mm/dd")
                    wRs3.Update
                End If
            Else
                    wRs3("実行日") = Format(Date, "yyyy/mm/dd")
                    wRs3.Update
            End If
        Else
            MsgBox "DAAA000_バージョンがありません"
                
            wRs3.Close
            Set wRs3 = Nothing
            '----------< GSerDir KXXX.mdb Close >-----------------------------------
            wDb.Close
            Set wDb = Nothing
                    
            Exit Function
        End If
        wRs3.Close
        Set wRs3 = Nothing
    End If
    '
    '----------< GSerDir KXXX.mdb Close >-------------------------------------------
    wDb.Close
    Set wDb = Nothing
'

'
    '----------< Check EXE >--------------------------------------------------------
    ws01 = "": ws02 = ""
    If GCurDir <> GSerDir Then
        GRet = GAA100_EXEVER(ws01, ws02) '2005/02現在 製品名で判断
        If GRet <> True Then
            ws01 = GSerComputerName & " : " & ws01
            ws02 = GMyComputerName & " : " & ws02
            
            GRet = MsgBox("バージョンの違うEXEを参照しています。" & vbCrLf & ws01 & vbLf & ws02, vbOKOnly + vbCritical)
            Exit Function
        End If
    End If
'
    MAA100_MDBVER = True
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA100_MDBVER_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA100_MDBVER() でエラー" + vbCrLf + vbCrLf + _
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
' CHECK_SERDIR
'------------------------------------------------
Private Function CHECK_SERDIR() As Boolean
'
    Dim wl01 As Long
    Dim j As Integer, L As Integer
    Dim ws01 As String, ws02 As String
'
    On Error GoTo Err_Hundle
'
    CHECK_SERDIR = False
'
    ws01 = "": ws02 = ""
        
        ws01 = GSerDir + "\" + GMain
        ws02 = GSerDir + "\" + GTemp
        If Dir(ws01) = "" Or Dir(ws02) = "" Then
            Exit Function
        End If
        
        '金剛石 or 借換たろう！
        If GProduct <> "金剛石" Then
            ws01 = GSerDir + "\" + "借換たろう.exe"
        Else
            ws01 = GSerDir + "\" + "金剛石.exe"
        End If
        If Dir(ws01) = "" Then
            Exit Function
        End If
        
        ws01 = "": ws02 = ""
        If GSerDir <> GCurDir Then
            If Left$(GSerDir, 2) Like "*:" Or GSerDir = "" Then
                GSerComputerName = GMyComputerName
            Else
                wl01 = Len(GSerDir)
                For j = 3 To wl01
                    ws01 = Mid$(GSerDir, j, 1)
                    If ws01 <> "\" Then
                        ws02 = ws02 & ws01
                    Else
                        Exit For
                    End If
                Next j
    
                GSerComputerName = ws02
            End If
        End If
    On Error GoTo 0
'
    CHECK_SERDIR = True
'
Exit Function
'----------< ERROR ROUTINE >--------------------------------------------------------
Err_Hundle:
    On Error GoTo 0
End Function

'------------------------------------------------
' GAA100_EXEVER
'------------------------------------------------
Public Function GAA100_EXEVER(pVer1 As String, pVer2 As String) As Boolean
'
    Dim cp As CODEPAGE
    Dim bBuffer() As Byte
    Dim lngRet As Long
    Dim lngDummy As Long
    Dim lngLen As Long
    Dim lpBuffer As Long
    Dim j As Integer
    
    Dim strPath As String
    Dim strFileName(1) As String
    Dim strProductName(1) As String
'
    On Error GoTo GAA100_EXEVER_ERR
'
    GAA100_EXEVER = False
'
    pVer1 = "": pVer2 = ""

    ' strFileName に取得したいファイル名をセット
    '金剛石 or 借換たろう！
    If GProduct <> "金剛石" Then
        strFileName(0) = GSerDir & "\" & "借換たろう.exe"
        strFileName(1) = GCurDir & "\" & "借換たろう.exe"
    Else
        strFileName(0) = GSerDir & "\" & "金剛石.exe"
        strFileName(1) = GCurDir & "\" & "金剛石.exe"
    End If

    For j = 0 To 1
        ' サイズを取得
        lngLen = GetFileVersionInfoSize(strFileName(j), lngDummy)
        If lngLen < 1 Then
            Exit Function
        End If
'
        ' バイトの配列の領域取得
        ReDim bBuffer(lngLen)

        ' ファイル バージョン情報を取得
        lngRet = GetFileVersionInfo(strFileName(j), 0&, lngLen, bBuffer(0))
        lngRet = VerQueryValue(bBuffer(0), "\VarFileInfo\Translation", lpBuffer, lngLen)

        ' 文字列情報の設定
        MoveMemory cp, lpBuffer, lngLen
        strPath = "\StringFileInfo\" & Right$("0000" & Hex$(cp.lngLOW), 4) & Right$("0000" & Hex$(cp.lngHIGH), 4) & "\"

        ' 製品名の取得
        lngRet = VerQueryValue(bBuffer(0), strPath & "ProductName", lpBuffer, lngLen)
        strProductName(j) = Space(lngLen)
        MoveMemory ByVal strProductName(j), lpBuffer, lngLen

        strProductName(j) = Left$(strProductName(j), 12)
    
    Next j
    
    If strProductName(0) <> strProductName(1) Then
        Exit Function
    End If
    
    pVer1 = strProductName(0)
    pVer2 = strProductName(1)
'
    GAA100_EXEVER = True
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
GAA100_EXEVER_ERR:
    pERR_MES = pPROGRAM_ID + "/ GAA100_EXEVER() でエラー" + vbCrLf + vbCrLf + _
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
' MAA100_KARIKAETAROU
'------------------------------------------------
Public Function MAA100_KARIKAETAROU() As Boolean
'
    Dim wJET As New JetEngine
    Dim wDb As New ADODB.Connection
    Dim wRs3 As ADODB.Recordset
    Dim wstr As String
'
    On Error GoTo MAA100_KARIKAETAROU_ERR
'
    MAA100_KARIKAETAROU = False
'
    ' =========================================
    '           有効日日付のＣＨＥＣＫ
    ' =========================================
    Call AdoDbOpen("Jet", wDb, GCurDir + "\" + GMain, "", , GPwd)
    
    wstr = "Select * From 借換たろうシステム制御"
    wstr = wstr + " Where System = 'System'"
    Call AdoRecordsetOpen(wDb, wRs3, wstr)
    If Not wRs3.EOF Then
        GDate1 = wRs3("終了日付")
    End If
    wRs3.Close
    Set wRs3 = Nothing

'    If Format(GDate1, "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
'        GRet = MsgBox("借換たろう！お試し版の有効日付が終了しました", vbOKOnly + vbCritical)
'        Exit Function
'    End If
'
    MAA100_KARIKAETAROU = True
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA100_KARIKAETAROU_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA100_KARIKAETAROU() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)
    
    End
'
End Function



