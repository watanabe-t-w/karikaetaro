Attribute VB_Name = "MXA040_Export"
Option Explicit
'
Private Const pPROGRAM_ID As String = "MXA040_Export"

Dim wRs As ADODB.Recordset
Dim wstr As String

Dim FLG_Check As Integer
Dim wCsvDir As String
Dim Nendo As String, kubun As String
Dim UriNo As String, KarNo As String, LeaNo As String, SetNo As String
Dim KinNo As String, SKinNo As String

Dim wdYushi(12) As Double, wdGankin(12) As Double, wdRisoku(12) As Double
Dim wdYushi2(12) As Double, wdGankin2(12) As Double, wdRisoku2(12) As Double
Dim wdHensai(12) As Double, wdKaiyaku(12) As Double, wdYZan(12) As Double
Dim wdHensai2(12) As Double, wdKaiyaku2(12) As Double, wdYZan2(12) As Double

Dim w推移表タイトル As MAA910_推移表タイトル
'
Dim w借入データ As MAA910_借入金
Dim w社債入力() As MAA910_借入金入力

Dim wField() As MXA040_TypeField
Private Type MXA040_TypeField
    Name As String
    Type As String
End Type
'
'------------------------------------------------
' MX040_CsvOut
'------------------------------------------------
Public Sub MX040_CsvOut()
'
    Dim ws01 As String
'
    On Error GoTo MX040_CsvOut_ERR
'
    FLG_Check = False
'
    Nendo = GRpt.テキスト_01
    kubun = GRpt.推移
    UriNo = "": KarNo = ""
    LeaNo = "": SetNo = ""
    KinNo = "": SKinNo = ""
'
    Select Case GRpt.帳票名
    Case "損益推移表"
        UriNo = GRpt.売上: KarNo = GRpt.借入
        LeaNo = GRpt.リス: SetNo = GRpt.設備
        KinNo = GRpt.金融: SKinNo = GRpt.設備R

        If SetNo <> "" Then
            Call MX040_Setubi(Gcsv_Set1)
        End If

        Call MX040_Kamoku(Gcsv_Shu1)
        Call MX040_KamTbl(Gcsv_Ktl1)

    Case "経営計画支援表"
        If GRpt.選択 = "1案" Or GRpt.選択 = "比較" Then
            UriNo = GRpt.売上: KarNo = GRpt.借入
            LeaNo = GRpt.リス: SetNo = GRpt.設備
            KinNo = GRpt.金融: SKinNo = GRpt.設備R

            If SetNo <> "" Then
                Call MX040_Setubi(Gcsv_Set1)
            End If

            Call MX040_Kamoku(Gcsv_Shu1)
            Call MX040_KamTbl(Gcsv_Ktl1)

        End If

        If GRpt.選択 = "2案" Or GRpt.選択 = "比較" Then
            UriNo = GRpt.売上2: KarNo = GRpt.借入2
            LeaNo = GRpt.リス2: SetNo = GRpt.設備2
            KinNo = GRpt.金融2: SKinNo = GRpt.設備R2

            If SetNo <> "" Then
                Call MX040_Setubi(Gcsv_Set2)
            End If

            Call MX040_Kamoku(Gcsv_Shu2)
            Call MX040_KamTbl(Gcsv_Ktl2)

        End If

    Case "分岐点売上表"
        UriNo = GRpt.売上: KarNo = GRpt.借入
        LeaNo = GRpt.リス: SetNo = GRpt.設備
        KinNo = GRpt.金融: SKinNo = GRpt.設備R

        If SetNo <> "" Then
            Call MX040_Setubi(Gcsv_Set1): Call MX040_Setubi(Gcsv_Set2)
        End If

        Call MX040_Kamoku(Gcsv_Shu1): Call MX040_Kamoku(Gcsv_Shu2)
        Call MX040_KamTbl(Gcsv_Ktl1): Call MX040_KamTbl(Gcsv_Ktl2)

    Case "損益予実対比表"
        kubun = "月次"
        UriNo = GRpt.売上: KarNo = GRpt.借入
        LeaNo = GRpt.リス: SetNo = GRpt.設備
        KinNo = GRpt.金融: SKinNo = GRpt.設備R

        If SetNo <> "" Then
            Call MX040_Setubi(Gcsv_Set1): Call MX040_Setubi(Gcsv_Set2)
        End If

        Call MX040_Kamoku(Gcsv_Shu1): Call MX040_Kamoku(Gcsv_Shu2)
        Call MX040_KamTbl(Gcsv_Ktl1): Call MX040_KamTbl(Gcsv_Ktl2)

    Case "固定資産台帳"
        Call MX040_固定資産台帳("固定資産台帳.csv")

'    Case "設備計画推移表"
'        kubun = "月次"
'
'        Call MX040_Setubi_SM("設備計画推移表.csv")

    Case Else
        Exit Sub
    End Select
'
    If FLG_Check = True Then
       'MsgBox "対象ファイル起動中の為、CSVファイルを作成できませんでした。", vbOKOnly + vbInformation
    End If
    
    FLG_Check = False
'    Screen.MousePointer = Default
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MX040_CsvOut_ERR:
    pERR_MES = pPROGRAM_ID + "/ MX040_CsvOut() でエラー" + vbCrLf + vbCrLf + _
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
' MX040_Kamoku
'------------------------------------------------
Private Sub MX040_Kamoku(pCsvFileName As String)
'
    Dim j As Integer
    Dim wTbl_sh As String, ws01 As String
'
    On Error Resume Next
'
    'Call MX040_CrtShemaini
    
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    'Select DataTable
    If pCsvFileName = Gcsv_Shu1 Then
        wTbl_sh = "DECA010_科目集計"
    ElseIf pCsvFileName = Gcsv_Shu2 Then
        wTbl_sh = "DECA020_科目集計２"
    End If
    
    '売上計画番号
    If Len(UriNo) = 2 Then
        wstr = "Select K.科目 As 科目番号,"
        wstr = wstr & "M.科目名,"
        
        'Parameter
        wstr = wstr & "'" & G基本情報.決算月 & "' As 決算月,"
        wstr = wstr & "'" & Nendo & "' As 推移表開始年度,"
        wstr = wstr & "'" & kubun & "' As 推移表区分,"
        wstr = wstr & "'" & UriNo & "' As 売上計画番号,"
        wstr = wstr & "'' As 売上計画内容,"
        wstr = wstr & "'" & KarNo & "' As 借入計画番号,"
        wstr = wstr & "'" & LeaNo & "' As リース計画番号,"
        wstr = wstr & "'" & SetNo & "' As 設備計画番号,"
        wstr = wstr & "'" & KinNo & "' As 金融リストラ番号,"
        wstr = wstr & "'" & SKinNo & "' As 設備リストラ番号,"
        wstr = wstr & "'" & GCoName & "' As 企業名,"
        
        wstr = wstr & "K.金額合計,"
        For j = 1 To 11
            wstr = wstr & "K.金額" & CStr(j) & "番目,"
        Next j
        wstr = wstr & "K.金額12番目"
        
        wstr = wstr & " INTO [Text;database=" & wCsvDir & "]." & "[" & pCsvFileName & "]"
        wstr = wstr & " FROM " & wTbl_sh & " As K"
        wstr = wstr & " INNER JOIN DAAA030_科目マスタ As M"
        wstr = wstr & " ON K.科目 = M.科目番号"
    Else
        wstr = "Select K.科目 As 科目番号,"
        wstr = wstr & "M.科目名,"
        
        'Parameter
        wstr = wstr & "'" & G基本情報.決算月 & "' As 決算月,"
        wstr = wstr & "'" & Nendo & "' As 推移表開始年度,"
        wstr = wstr & "'" & kubun & "' As 推移表区分,"
        wstr = wstr & "'" & UriNo & "' As 売上計画番号,"
        wstr = wstr & "U.売上計画内容,"
        wstr = wstr & "'" & KarNo & "' As 借入計画番号,"
        wstr = wstr & "'" & LeaNo & "' As リース計画番号,"
        wstr = wstr & "'" & SetNo & "' As 設備計画番号,"
        wstr = wstr & "'" & KinNo & "' As 金融リストラ番号,"
        wstr = wstr & "'" & SKinNo & "' As 設備リストラ番号,"
        wstr = wstr & "'" & GCoName & "' As 企業名,"
        
        wstr = wstr & "K.金額合計,"
        For j = 1 To 11
            wstr = wstr & "K.金額" & CStr(j) & "番目,"
        Next j
        wstr = wstr & "K.金額12番目"
        
        wstr = wstr & " INTO [Text;database=" & wCsvDir & "]." & "[" & pCsvFileName & "]"
        wstr = wstr & " FROM DBAA040_売上計画 As U," & wTbl_sh & " As K"
        wstr = wstr & " INNER JOIN DAAA030_科目マスタ As M"
        wstr = wstr & " ON K.科目 = M.科目番号"
        wstr = wstr & " Where U.売上計画番号 = " & "'" & UriNo & "'"
    End If
    
    GDb.Execute wstr
    
    DoEvents
    
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
End Sub

'------------------------------------------------
' MX040_KamTbl
'------------------------------------------------
Private Sub MX040_KamTbl(pCsvFileName As String)
'
    Dim j As Integer
    Dim wTbl_sh As String, ws01 As String
'
    On Error Resume Next
'
    'Call MX040_CrtShemaini
    
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    'Select DataTable
    If pCsvFileName = Gcsv_Ktl1 Then
        wTbl_sh = "DXAA020_科目テーブルデバック"
    ElseIf pCsvFileName = Gcsv_Ktl2 Then
        wTbl_sh = "DXAA020_科目テーブルデバック２"
    End If
    
    '売上計画番号
    If Len(UriNo) = 2 Then
        wstr = "SELECT "
        wstr = wstr & "K.科目番号,"
        wstr = wstr & "M.科目名,"
        
        'Parameter
        wstr = wstr & "'" & G基本情報.決算月 & "' As 決算月,"
        wstr = wstr & "'" & Nendo & "' As 推移表開始年度,"
        wstr = wstr & "'" & kubun & "' As 推移表区分,"
        wstr = wstr & "'" & UriNo & "' As 売上計画番号,"
        wstr = wstr & "'' As 売上計画内容,"
        wstr = wstr & "'" & KarNo & "' As 借入計画番号,"
        wstr = wstr & "'" & LeaNo & "' As リース計画番号,"
        wstr = wstr & "'" & SetNo & "' As 設備計画番号,"
        wstr = wstr & "'" & KinNo & "' As 金融リストラ番号,"
        wstr = wstr & "'" & SKinNo & "' As 設備リストラ番号,"
        wstr = wstr & "'" & GCoName & "' As 企業名,"
        
        For j = 1 To 131
            wstr = wstr & "K.数値" & CStr(j) & "番目,"
        Next j
        wstr = wstr & "K.数値132番目"
        
        wstr = wstr & " INTO [Text;database=" & wCsvDir & "]." & "[" & pCsvFileName & "]"
        wstr = wstr & " FROM " & wTbl_sh & " As K"
        wstr = wstr & " INNER JOIN DAAA030_科目マスタ As M"
        wstr = wstr & " ON K.科目番号 = M.科目番号"
    Else
        wstr = "SELECT "
        wstr = wstr & "K.科目番号,"
        wstr = wstr & "M.科目名,"
        
        'Parameter
        wstr = wstr & "'" & G基本情報.決算月 & "' As 決算月,"
        wstr = wstr & "'" & Nendo & "' As 推移表開始年度,"
        wstr = wstr & "'" & kubun & "' As 推移表区分,"
        wstr = wstr & "'" & UriNo & "' As 売上計画番号,"
        wstr = wstr & "U.売上計画内容,"
        wstr = wstr & "'" & KarNo & "' As 借入計画番号,"
        wstr = wstr & "'" & LeaNo & "' As リース計画番号,"
        wstr = wstr & "'" & SetNo & "' As 設備計画番号,"
        wstr = wstr & "'" & KinNo & "' As 金融リストラ番号,"
        wstr = wstr & "'" & SKinNo & "' As 設備リストラ番号,"
        wstr = wstr & "'" & GCoName & "' As 企業名,"
        
        For j = 1 To 131
            wstr = wstr & "K.数値" & CStr(j) & "番目,"
        Next j
        wstr = wstr & "K.数値132番目"
        
        wstr = wstr & " INTO [Text;database=" & wCsvDir & "]." & "[" & pCsvFileName & "]"
        wstr = wstr & " FROM DBAA040_売上計画 As U," & wTbl_sh & " As K"
        wstr = wstr & " INNER JOIN DAAA030_科目マスタ As M"
        wstr = wstr & " ON K.科目番号 = M.科目番号"
        wstr = wstr & " Where U.売上計画番号 = " & "'" & UriNo & "'"
    End If
    
    GDb.Execute wstr
    
    DoEvents
    
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
End Sub

'------------------------------------------------
' MX040_Setubi
'------------------------------------------------
Private Sub MX040_Setubi(pCsvFileName As String)
'
    On Error Resume Next
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    wstr = "Select "
    wstr = wstr & "S.設備番号,S.設備名,"
    wstr = wstr & "S.設備計画番号,"
    wstr = wstr & "IIF(S.sm区分=1,'シミュレーション','') As SM区分,"
'    wstr = wstr & "S.残高移行区分,"
    wstr = wstr & "S.設備年度,"
    wstr = wstr & "Format(S.設備年月,'yyyy/mm/dd') As 設備年月,"
    wstr = wstr & "Format(S.設備購入年月日,'yyyy/mm/dd') As 設備購入年月日,"
    wstr = wstr & "Format(S.償却最終年月,'yyyy/mm/dd') As 償却最終年月,"
    wstr = wstr & "IIF(S.償却区分='" & XMXA020_区分("償却区分", "定率法") & "',M.定率法償却率,M.定額法償却率) As 償却率,"
    wstr = wstr & "IIF(S.償却区分='" & XMXA020_区分("償却区分", "定額法") & "','定額法',"
    wstr = wstr & " IIF(S.償却区分='" & XMXA020_区分("償却区分", "定率法") & "','定率法',"
    wstr = wstr & " IIF(S.償却区分='" & XMXA020_区分("償却区分", "均等償却") & "','均等償却',''))) As 償却区分,"
    wstr = wstr & "IIF(S.資産区分='" & XMXA020_区分("資産区分", "有形資産") & "','有形資産',"
    wstr = wstr & " IIF(S.資産区分='" & XMXA020_区分("資産区分", "無形資産") & "','無形資産',"
    wstr = wstr & " IIF(S.資産区分='" & XMXA020_区分("資産区分", "損金設備") & "','損金設備',"
    wstr = wstr & " IIF(S.資産区分='" & XMXA020_区分("資産区分", "建物") & "','建物',"
    wstr = wstr & " IIF(S.資産区分='" & XMXA020_区分("資産区分", "土地") & "','土地',''))))) As 資産区分,"
    wstr = wstr & "S.設備金額,"
    wstr = wstr & "IIF(S.課税区分='" & XMXA020_区分("課税区分", "不課税") & "','不課税',"
    wstr = wstr & " IIF(S.課税区分='" & XMXA020_区分("課税区分", "課税") & "','課税',"
    wstr = wstr & " IIF(S.課税区分='" & XMXA020_区分("課税区分", "課税") & "','非課税',''))) As 購入課税区分,"
    wstr = wstr & "S.支払サイト,"
    wstr = wstr & "S.償却年数,S.残存率,"
    
    wstr = wstr & "S.調整償却額,S.特別償却１年次額,S.特別償却２年次額,S.特別償却３年次額,"
    
    wstr = wstr & "S.設備リストラ番号,"
    wstr = wstr & "Format(S.資産売却年月日,'yyyy/mm/dd') As 資産売却年月日,"
    wstr = wstr & "S.資産売却額,"
    wstr = wstr & "IIF(S.売上課税区分='" & XMXA020_区分("課税区分", "不課税") & "','不課税',"
    wstr = wstr & "IIF(S.売上課税区分='" & XMXA020_区分("課税区分", "課税") & "','課税',"
    wstr = wstr & "IIF(S.売上課税区分='" & XMXA020_区分("課税区分", "課税") & "','非課税',''))) As 売却課税区分,"
    wstr = wstr & "S.回収サイト,"
    wstr = wstr & "IIF(S.手入力フラグ=0,'基幹データ','') As データ区分"
'    wstr = wstr & "S.修正不可F"
    
    wstr = wstr & " INTO [Text;database=" & wCsvDir & "]." & "[" & pCsvFileName & "]"
    wstr = wstr & " From (DBCA010_設備計画  As S"
    wstr = wstr & " Left Join DAAB010_償却率マスタ As M"
    wstr = wstr & " ON S.償却年数 = M.償却年数)"
    wstr = wstr & " Where S.設備計画番号 = '" & SetNo & "'"
    wstr = wstr & " AND S.設備リストラ番号 = '" & SKinNo & "'"
    wstr = wstr & " AND 取消フラグ = 0 "
    wstr = wstr & " Order BY S.設備年月,S.設備購入年月日"
    
    GDb.Execute wstr
    
    DoEvents
    
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
End Sub

'------------------------------------------------
' MX040_Setubi_SM
'------------------------------------------------
Private Sub MX040_Setubi_SM(pCsvFileName As String)
'
    Dim ws01 As String
    Dim wsNendo As String
'
    On Error Resume Next
'
    'Call MX040_CrtShemaini
    
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wsNendo = P8.FCDbl(C年月日.年度開始年月日(Nendo, "y"))
'
    wstr = "SELECT "
    wstr = wstr & " IIf(Left(S.設備番号,4)='" & wsNendo & "','設備2','設備1') AS データ区分,"
    wstr = wstr & " S.資産区分,"
    wstr = wstr & " S.償却区分,"
    wstr = wstr & " Count(K.設備番号) AS 件数,"
    wstr = wstr & " Sum(K.新規_01) AS 新規合計,"
    wstr = wstr & " Sum(K.期首_01) AS 期首合計,"
    wstr = wstr & " Sum(K.償却_01) AS 償却合計,"
    wstr = wstr & " Sum(K.調整償却_01) AS 調整償却合計,"
    wstr = wstr & " Sum(K.特別償却_01) AS 特別償却合計,"
    wstr = wstr & " Sum(K.売却額_01) AS 売却額合計,"
    wstr = wstr & " Sum(K.売却益_01) AS 売却益合計,"
    wstr = wstr & " Sum(K.売却損_01) AS 売却損合計,"
    wstr = wstr & " Sum(K.残存_01) AS 残存合計"
    wstr = wstr & " INTO [Text;database=" & wCsvDir & "]." & "[" & pCsvFileName & "]"
    wstr = wstr & " FROM DCCA010_設備推移結果 AS K"
    wstr = wstr & " INNER JOIN DBCA010_設備計画 AS S"
    wstr = wstr & " ON K.設備番号 = S.設備番号"
    wstr = wstr & " GROUP BY IIf(Left(S.設備番号,4)='" & wsNendo & "','設備2','設備1'), S.資産区分, S.償却区分"
    wstr = wstr & " ORDER BY IIf(Left(S.設備番号,4)='" & wsNendo & "','設備2','設備1'), S.資産区分, S.償却区分"
    
    GDb.Execute wstr
    
    DoEvents
    
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
End Sub

'------------------------------------------------
' MX040_CsvPath
'------------------------------------------------
'Public Sub MX040_CsvPath()
''
'    Dim objFileSystem As Object
'    Dim objFile As Object
'    Dim ws01 As String
''
'    ws01 = GCurDir & "\" & GTemp
'    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
'    Set objFile = objFileSystem.GetFile(ws01)
'        ws01 = UCase(objFile.Drive)
'    Set objFile = Nothing
'    Set objFileSystem = Nothing
'
'    wCsvDir = ws01 & "\" & Gcsv_DirName
''
'End Sub
'------------------------------------------------
' MX040_CsvPath
'------------------------------------------------
Public Sub MX040_CsvPath()
'
    Dim objFileSystem As Object
    Dim objFile As Object
    Dim ws01 As String

    '@001 ADD STR CSV出力先
    Dim wDb As New ADODB.Connection
    Dim wRs3 As ADODB.Recordset
'
    If GCsvPath = "" Then
        '----------< GSerDir GMain Open >-----------------------------------------------
        Call AdoDbOpen("Jet", wDb, GCurDir + "\" + GMain, "", , GPwd)
            
        wstr = "Select * From LIST000_データ保存先マスタ"
        wstr = wstr + " Where System = 'System'"
        Call AdoRecordsetOpen(wDb, wRs3, wstr)
        If Not wRs3.eof Then
            GCsvPath = P8.FCStr(wRs3("CSVPATH"))
        End If
        wRs3.Close
        Set wRs3 = Nothing
        
        '----------< GSerDir GMain Close >----------------------------------------------
        wDb.Close
        Set wDb = Nothing
        
    End If
'
    'Check CsvPath
    If GCsvPath <> "" Then
        On Error Resume Next
        
        ws01 = wCsvDir
        If Dir(GCsvPath, vbDirectory) <> "" Then
            If Err.Number Then
                Err.Clear
            Else
                wCsvDir = GCsvPath
                Exit Sub
            End If
        End If
    
        On Error GoTo 0
    End If
    '@001 ADND
'
    ws01 = GCurDir & "\" & GTemp
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFileSystem.GetFile(ws01)
        ws01 = UCase(objFile.Drive)
    Set objFile = Nothing
    Set objFileSystem = Nothing
    
    wCsvDir = ws01 & "\" & Gcsv_DirName
'
End Sub

'------------------------------------------------
' MX040_CsvPath
'------------------------------------------------
Public Sub MX040_CsvPath_CDL(pPath As String)
'
    wCsvDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(pPath)
'
End Sub

'------------------------------------------------
' MX040_CsvFile_Check
'------------------------------------------------
Private Function MX040_CsvFile_Check(pCsvFileName As String) As Boolean
'
    Dim hWnd As Long
    Dim ws01 As String
    Dim st As Long
'
    On Error Resume Next
'
    MX040_CsvFile_Check = False
    
'----------< Create \金剛石CSV or \借換たろうCSV >----------------------------------
    ws01 = wCsvDir
    If Dir(ws01, vbDirectory) = "" Then
        MkDir (ws01)
        
        If Err.Number Then
            Err.Clear
            Exit Function
        End If
        
        MX040_CsvFile_Check = True
        Exit Function
    End If

    st = Timer
    Do While Timer - st < 2   '2秒間待つ
        DoEvents
    Loop
'
'----------< 指定のファイルが使用中かどうかを調べる >---------------------
    
    '起動中ならハンドルが返り、起動していなければ0が返る
    'Hwnd = FindWindow("XLMAIN", vbNullString)      'クラス名を与えて エクセルのハンドルを取得
    
    ws01 = "Microsoft Excel - " & pCsvFileName      '参考キャプション名を与えてハンドルを取得
    hWnd = FindWindow(vbNullString, ws01)
    If hWnd <> 0 Then
        Exit Function
'        GRet = MsgBox("CSVファイル作成の為、起動中のExcelを終了します。", vbYesNo)
'        If GRet = vbNo Then
'            Exit Function
'        Else
            'エクセル（アプリ）が起動中なら終了する場合
'            GRet = SendMessage(Hwnd, WM_CLOSE, 0&, 0&)'指定のハンドルに終了のメッセージを送る
'        End If
    End If
    
    'ファイルがあればファイルの名前を同じ名前で変更します。ファイルが使用中であればエラーが発生します。
    ws01 = wCsvDir & "\" & pCsvFileName
    If Dir(ws01) <> "" Then
    
        Name ws01 As ws01
        
        If Err.Number Then
            Err.Clear
            Exit Function
        End If
    End If
    
    'CsvFile Dlete
    ws01 = wCsvDir & "\" & pCsvFileName
    If Dir(ws01) <> "" Then
        Kill (ws01)
        
        If Err.Number Then
            Err.Clear
            Exit Function
        End If
    End If
'
    MX040_CsvFile_Check = True
'
    On Error GoTo 0
'
End Function

'------------------------------------------------
' MX040_CrtShemaini
'------------------------------------------------
Public Sub MX040_CrtShemaini()
'
    Dim I As Integer, j As Integer, k As Integer, l As Integer
    Dim intFileno As Integer, wi01 As Integer
    Dim ws01 As String
    Dim ws_shu(1) As String, ws_Ktl(1) As String, ws_Set(1) As String
'
    On Error GoTo MX040_CrtShemaini_ERR
'
'----------< Initialize >-----------------------------------------------------------
    ws_shu(0) = Gcsv_Shu1: ws_shu(1) = Gcsv_Shu2
    ws_Ktl(0) = Gcsv_Ktl1: ws_Ktl(1) = Gcsv_Ktl2
    ws_Set(0) = Gcsv_Set1: ws_Set(1) = Gcsv_Set2
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If

'----------< Create \金剛石CSV or \借換たろうCSV >----------------------------------
    ws01 = wCsvDir
    If Dir(ws01, vbDirectory) = "" Then
        MkDir (ws01)
        
        If Err.Number Then
            Err.Clear
            Exit Sub
        End If
    End If
'
    DoEvents
'
'----------< schema.ini Create >----------------------------------------------------
    intFileno = FreeFile
    Open wCsvDir & "\schema.ini" For Output As intFileno
'
    'Kcsv_科目集計
    For l = 0 To 1
        Print #intFileno, "[" & ws_shu(l) & "]"
        Print #intFileno, "ColNameHeader = True"
        Print #intFileno, "CharacterSet = OEM"
        Print #intFileno, "Format = CSVDelimited"
        
        Print #intFileno, "Col1=科目番号 Char Width 10"
        Print #intFileno, "Col2=科目名 Char Width 20"
        Print #intFileno, "Col3=決算月 Short"
        Print #intFileno, "Col4=推移表開始年度 Short"
        Print #intFileno, "Col5=推移表区分 Char Width 20"
        Print #intFileno, "Col6=連結売上計画番号 Char Width 20"
        Print #intFileno, "Col7=売上計画番号 Char Width 20"
        Print #intFileno, "Col8=売上計画内容 Char Width 50"
        Print #intFileno, "Col9=借入計画番号 Char Width 20"
        Print #intFileno, "Col10=リース計画番号 Char Width 20"
        Print #intFileno, "Col11=設備計画番号 Char Width 20"
        Print #intFileno, "Col12=金融リストラ番号 Char Width 20"
        Print #intFileno, "Col13=設備リストラ番号 Char Width 20"
        Print #intFileno, "Col14=企業名 Char Width 50"
        Print #intFileno, "Col15=金額合計 Float"
        
        wi01 = 16
        For k = 1 To 12
            ws01 = "Col" & wi01 & "=金額" & CStr(k) & "番目 Float"
            Print #intFileno, ws01
            
            wi01 = wi01 + 1
        Next k
    Next l
'
    'Kcsv_科目テーブル
    For l = 0 To 1
        Print #intFileno, "[" & ws_Ktl(l) & "]"
        Print #intFileno, "ColNameHeader = True"
        Print #intFileno, "CharacterSet = OEM"
        Print #intFileno, "Format = CSVDelimited"
        
        Print #intFileno, "Col1=科目番号 Char Width 10"
        Print #intFileno, "Col2=科目名 Char Width 20"
        Print #intFileno, "Col3=決算月 Short"
        Print #intFileno, "Col4=推移表開始年度 Short"
        Print #intFileno, "Col5=推移表区分 Char Width 20"
        Print #intFileno, "Col6=連結売上計画番号 Char Width 20"
        Print #intFileno, "Col7=売上計画番号 Char Width 20"
        Print #intFileno, "Col8=売上計画内容 Char Width 50"
        Print #intFileno, "Col9=借入計画番号 Char Width 20"
        Print #intFileno, "Col10=リース計画番号 Char Width 20"
        Print #intFileno, "Col11=設備計画番号 Char Width 20"
        Print #intFileno, "Col12=金融リストラ番号 Char Width 20"
        Print #intFileno, "Col13=設備リストラ番号 Char Width 20"
        Print #intFileno, "Col14=企業名 Char Width 50"
        
        wi01 = 15
        For k = 1 To 132
            ws01 = "Col" & wi01 & "=数値" & CStr(k) & "番目 Float"
            Print #intFileno, ws01
        
            wi01 = wi01 + 1
        Next k
    Next l
'
    'Kcsv_設備計画
    For l = 0 To 1
        Print #intFileno, "[" & ws_Set(l) & "]"
        Print #intFileno, "ColNameHeader = True"
        Print #intFileno, "CharacterSet = OEM"
        Print #intFileno, "Format = CSVDelimited"
        
        Print #intFileno, "Col1=設備計画番号 Char Width 20"
        Print #intFileno, "Col2=設備番号 Char Width 20"
        Print #intFileno, "Col3=設備名 Char Width 30"
        Print #intFileno, "Col4=設備年度 Short"
        Print #intFileno, "Col5=設備年月 Date"
        Print #intFileno, "Col6=償却最終年月 Date"
        Print #intFileno, "Col7=設備金額 Float"
        Print #intFileno, "Col8=償却年数 Short"
        Print #intFileno, "Col9=償却率 Float"
        Print #intFileno, "Col10=償却区分 Char Width 20"
        Print #intFileno, "Col11=資産区分 Char Width 20"
        Print #intFileno, "Col12=課税区分 Char Width 20"
    Next l
'
    Close #intFileno
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MX040_CrtShemaini_ERR:
    pERR_MES = pPROGRAM_ID + "/ MX040_CrtShemaini() でエラー" + vbCrLf + vbCrLf + _
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
' MX040_CsvOut2
'------------------------------------------------
Public Sub MX040_CsvOut2()
'
    Dim ws01 As String
'
    On Error GoTo MX040_CsvOut2_ERR
'
    FLG_Check = False
'
    Nendo = GRpt.テキスト_01
    kubun = GRpt.推移
    UriNo = "": KarNo = ""
    LeaNo = "": SetNo = ""
    KinNo = "": SKinNo = ""
'
    UriNo = GRpt.売上: KarNo = GRpt.借入
    LeaNo = GRpt.リス: SetNo = GRpt.設備
    KinNo = GRpt.金融: SKinNo = GRpt.設備R
'
    'デモ用
    FLG_Check = True
    
    Exit Sub
'
    Call MX040_Kamoku2(Gcsv_Shu1)
'
    If FLG_Check = True Then
       'MsgBox "対象ファイル起動中の為、CSVファイルを作成できませんでした。", vbOKOnly + vbInformation
    End If
    
    FLG_Check = False
'    Screen.MousePointer = Default
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MX040_CsvOut2_ERR:
    pERR_MES = pPROGRAM_ID + "/ MX040_CsvOut2() でエラー" + vbCrLf + vbCrLf + _
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
' MX040_Kamoku2
'------------------------------------------------
Private Sub MX040_Kamoku2(pCsvFileName As String)
'
    Dim j As Integer
    Dim wTbl_sh As String, ws01 As String
'
    On Error Resume Next
'
    'Call MX040_CrtShemaini
    
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    'Select DataTable
    wTbl_sh = "DCXA020_帳票作成ワーク"
    
    '売上計画番号
    If Len(UriNo) = 2 Then
        wstr = "Select K.科目番号,"
        wstr = wstr & "K.科目名,"
        
        'Parameter
        wstr = wstr & "'" & G基本情報.決算月 & "' As 決算月,"
        wstr = wstr & "'" & Nendo & "' As 推移表開始年度,"
        wstr = wstr & "'" & kubun & "' As 推移表区分,"
        wstr = wstr & "'" & UriNo & "' As 計画番号,"
        wstr = wstr & "'基本事業計画' As 売上計画内容,"
        wstr = wstr & "IIF(K.売上計画番号='10','期累計','" & GRpt.推移 & "') As 期区分,"
        wstr = wstr & "'" & KarNo & "' As 借入計画番号,"
        wstr = wstr & "'" & LeaNo & "' As リース計画番号,"
        wstr = wstr & "'" & SetNo & "' As 設備計画番号,"
        wstr = wstr & "'" & KinNo & "' As 金融リストラ番号,"
        wstr = wstr & "'" & SKinNo & "' As 設備リストラ番号,"
        wstr = wstr & "'" & GCoName & "' As 企業名,"
        
        For j = 1 To 19
            wstr = wstr & "K.コード_" & Right$("000" & CStr(j), 3) & ","
        Next j
        wstr = wstr & "K.コード_020"
        
        wstr = wstr & " INTO [Text;database=" & wCsvDir & "]." & "[" & pCsvFileName & "]"
        wstr = wstr & " FROM " & wTbl_sh & " As K"
        wstr = wstr & " INNER JOIN DAAA030_科目マスタ As M"
        wstr = wstr & " ON K.科目番号 = M.科目番号"
    
    Else
        wstr = "Select K.科目番号,"
        wstr = wstr & "K.科目名,"
        
        'Parameter
        wstr = wstr & "'" & G基本情報.決算月 & "' As 決算月,"
        wstr = wstr & "'" & Nendo & "' As 推移表開始年度,"
        wstr = wstr & "'" & kubun & "' As 推移表区分,"
        wstr = wstr & "'" & UriNo & "' As 計画番号,"
        wstr = wstr & "U.売上計画内容,"
        wstr = wstr & "IIF(K.売上計画番号='10','期累計','" & GRpt.推移 & "') As 期区分,"
        wstr = wstr & "'" & KarNo & "' As 借入計画番号,"
        wstr = wstr & "'" & LeaNo & "' As リース計画番号,"
        wstr = wstr & "'" & SetNo & "' As 設備計画番号,"
        wstr = wstr & "'" & KinNo & "' As 金融リストラ番号,"
        wstr = wstr & "'" & SKinNo & "' As 設備リストラ番号,"
        wstr = wstr & "'" & GCoName & "' As 企業名,"
        
        For j = 1 To 19
            wstr = wstr & "K.コード_" & Right$("000" & CStr(j), 3) & ","
        Next j
        wstr = wstr & "K.コード_020"
        
        wstr = wstr & " INTO [Text;database=" & wCsvDir & "]." & "[" & pCsvFileName & "]"
        wstr = wstr & " FROM DBAA040_売上計画 As U," & wTbl_sh & " As K"
        wstr = wstr & " INNER JOIN DAAA030_科目マスタ As M"
        wstr = wstr & " ON K.科目番号 = M.科目番号"
        wstr = wstr & " Where U.売上計画番号 = " & "'" & UriNo & "'"
    
    End If
    
    GDb.Execute wstr
    
    DoEvents
    
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
End Sub

'------------------------------------------------
' MX040_CsvOut_KARI
'------------------------------------------------
Public Sub MX040_CsvOut_KARI()
'
    Dim ws01 As String
    Dim wsFD1 As String, wsFD2 As String
'
    On Error GoTo MX040_CsvOut_KARI_ERR
'
    '借入金 Lite
    If GSys.Sys = "借入金 Lite" Then
        Exit Sub
    End If
'
    FLG_Check = False
'
    Nendo = GRpt.テキスト_01
    kubun = GRpt.推移
    UriNo = "": KarNo = ""
    LeaNo = "": SetNo = ""
    KinNo = "": SKinNo = ""
'
    Select Case GRpt.帳票名
    Case "借入金台帳"
        Call MX040_借入金台帳(GKeyName & "_" & GRpt.コンボ_01 & "借入金台帳.csv")
    
    Case "借入明細表"
        Call MX040_借入明細表(GKeyName & "_" & GRpt.コンボ_01 & "借入明細表.csv")
    
    Case "社債明細表"
        Call MX040_社債明細表(GKeyName & "_" & GRpt.コンボ_01 & "社債明細表.csv")
    
    Case "借入一覧表"
        Call MX040_借入一覧表(GKeyName & "_" & "借入一覧表.csv")
    
    Case "借入一覧表(全件)"
        Call MX040_借入一覧表_全件(GKeyName & "_" & "借入一覧表_全件.csv")
    
    Case "貸付一覧表"
        Call MX040_借入一覧表(GKeyName & "_" & "貸付一覧表.csv")
    
    Case "貸付明細表"
        Call MX040_借入明細表(GKeyName & "_" & GRpt.コンボ_01 & "貸付明細表.csv")
    
    Case "借入金返済予定表"
        Call MX040_借入金返済予定表(GKeyName & "_" & "借入金返済予定表.csv")
    
    Case "利息明細表"
        Call MX040_利息前払未払明細表(GKeyName & "_" & GRpt.コンボ_01 & "利息明細表.csv")
    
    Case "仕訳表"

        If G基本情報.日付入力区分 = "0" Then
        '和暦入力
            If GRpt.テキスト_01 <> GRpt.テキスト_02 Then
                ws01 = GRpt.テキスト_01 & "～" & GRpt.テキスト_02
            Else
                ws01 = GRpt.テキスト_01
            End If
        Else
        '西暦入力
            If GRpt.テキスト_01 <> GRpt.テキスト_02 Then
                wsFD1 = Format(GRpt.テキスト_01, "yyyymm")
                wsFD2 = Format(GRpt.テキスト_02, "yyyymm")
                ws01 = wsFD1 & "～" & wsFD2
            Else
                wsFD1 = Format(GRpt.テキスト_01, "yyyymm")
                ws01 = wsFD1
            End If
        End If

        Call MX040_仕訳表(GKeyName & "_" & ws01 & "仕訳表.csv")

    Case "仕訳表 -月次処理-"
        If G基本情報.日付入力区分 = "0" Then
        '和暦入力
            If GRpt.テキスト_01 <> GRpt.テキスト_02 Then
                ws01 = GRpt.テキスト_01 & "～" & GRpt.テキスト_02
            Else
                ws01 = GRpt.テキスト_01
            End If
        Else
        '西暦入力
            If GRpt.テキスト_01 <> GRpt.テキスト_02 Then
                wsFD1 = Format(GRpt.テキスト_01, "yyyymm")
                wsFD2 = Format(GRpt.テキスト_02, "yyyymm")
                ws01 = wsFD1 & "～" & wsFD2
            Else
                wsFD1 = Format(GRpt.テキスト_01, "yyyymm")
                ws01 = wsFD1
            End If
        End If

        '神姫バス
        'Call MX040_仕訳表_神姫バス(GKeyName & "_" & ws01 & "_仕訳表.csv")
        
        Call MX040_仕訳表(GKeyName & "_" & ws01 & "_仕訳表.csv")
        'Call MX040_仕訳表(GKeyName & "_" & "仕訳表.csv")
    
    Case "仕訳表 -決算処理-"
        If G基本情報.日付入力区分 = "0" Then
        '和暦入力
            ws01 = GRpt.テキスト_01
        Else
        '西暦入力
            ws01 = Format(GRpt.テキスト_01, "yyyymm")
        End If

        '神姫バス
        'Call MX040_仕訳表_神姫バス(GKeyName & "_" & ws01 & "_仕訳表_決算.csv")
        
        Call MX040_仕訳表(GKeyName & "_" & ws01 & "_仕訳表_決算.csv")
        'Call MX040_仕訳表(GKeyName & "_" & "仕訳表_決算.csv")
        
    Case "基準金利レート"
        Call MX040_長期プライムレート(GKeyName & "_" & GRpt.テキスト_01 & "_" & GRpt.帳票名 & ".csv")
        
    Case "借入金時価評価明細表"
        Call MX040_借入金時価評価明細表(GKeyName & "_" & GRpt.コンボ_01 & "借入金時価評価明細表.csv")
        
    Case "借入金時価評価適用金利一覧"
        Call MX040_借入金時価評価適用金利一覧(GKeyName & "_" & GCsvName & ".csv")
        
    Case "借入金時価評価適用金利一覧_前期末"
        Call MX040_借入金時価評価適用金利一覧(GKeyName & "_" & GCsvName & ".csv")
    
    Case "借入金時価評価一覧表"
        Call MX040_借入金時価評価一覧表(GKeyName & "_" & GCsvName & ".csv")

    Case "借入金時価評価一覧表_前期末"
        Call MX040_借入金時価評価一覧表(GKeyName & "_" & GCsvName & ".csv")

    Case "借入金時価評価一覧表_増減"
        Call MX040_借入金時価評価一覧表(GKeyName & "_" & GCsvName & ".csv")

    Case "長短振替表"
    '神姫バス
        Call MX040_長短振替表_神姫バス(GKeyName & "_" & "長短振替表.csv")
    
    Case "資金繰表"
    '神姫バス
        Call MX040_資金繰表_神姫バス(GKeyName & "_" & "資金繰表.csv")
    
    Case "1年以内返済長期借入金集計表"
    '杉村倉庫仕様
        ws01 = Format(GRpt.テキスト_02, "yyyymm")
        Call MX040_1年内返済集計表_杉村倉庫(GKeyName & "_" & ws01 & "_" & GRpt.帳票名 & ".csv")
    
    Case "銀行別利息表"
    '杉村倉庫仕様
        ws01 = Format(GRpt.テキスト_02, "yyyymm")
        Call MX040_銀行別利息表_杉村倉庫(GKeyName & "_" & ws01 & "_" & GRpt.帳票名 & "_利息先払.csv", "利息先払")
        Call MX040_銀行別利息表_杉村倉庫(GKeyName & "_" & ws01 & "_" & GRpt.帳票名 & "_利息後払.csv", "利息後払")

    Case "支払利息推移表"
    '杉村倉庫仕様
        Call MX040_支払利息推移表_杉村倉庫(GKeyName & "_" & GRpt.テキスト_01 & "年度" & Left(GRpt.テキスト_02, 5) & "_" & GRpt.帳票名 & ".csv")

    Case Else
        Exit Sub
    End Select
'
    If FLG_Check = True Then
       'MsgBox "対象ファイル起動中の為、CSVファイルを作成できませんでした。", vbOKOnly + vbInformation
    End If
    
    FLG_Check = False
'    Screen.MousePointer = Default
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MX040_CsvOut_KARI_ERR:
    pERR_MES = pPROGRAM_ID + "/ MX040_CsvOut_KARI() でエラー" + vbCrLf + vbCrLf + _
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
' MX040_CsvOut_KARISUII
'------------------------------------------------
Public Sub MX040_CsvOut_KARISUII(p推移表タイトル As MAA910_推移表タイトル)
'
    Dim ws01 As String
    Dim wsFD As String
'
    On Error GoTo MX040_CsvOut_KARISUII_ERR
'
    '借入金 Lite
    If GSys.Sys = "借入金 Lite" Then
        Exit Sub
    End If
'
    FLG_Check = False
'
    w推移表タイトル = p推移表タイトル

    Nendo = GRpt.テキスト_01
    kubun = GRpt.推移
    UriNo = "": KarNo = ""
    LeaNo = "": SetNo = ""
    KinNo = "": SKinNo = ""
'
    Select Case GRpt.帳票名
    Case "借入残高推移表"
    
        kubun = "月次"
        KinNo = GRpt.金融
        
        Call MX040_借入残高推移表(GKeyName & "_" & "借入残高推移表.csv")
    
    Case "貸付残高推移表"
    
        kubun = "月次"
        KinNo = GRpt.金融
        
        Call MX040_借入残高推移表(GKeyName & "_" & "貸付残高推移表.csv")
    
    Case "社債残高推移表"
    
        kubun = "月次"
        KinNo = GRpt.金融
        
        Call MX040_社債残高推移表(GKeyName & "_" & "社債残高推移表.csv")
    
    Case "借入利息残高推移表"
    
        kubun = "月次"
        KinNo = GRpt.金融
        
        Call MX040_前払利息残高推移表(GKeyName & "_" & "前払利息残高推移表.csv")
        DoEvents
        Call MX040_未払利息残高推移表(GKeyName & "_" & "未払利息残高推移表.csv")
    
    Case "貸付利息残高推移表"
    
        kubun = "月次"
        KinNo = GRpt.金融
        
        Call MX040_前払利息残高推移表(GKeyName & "_" & "貸付前払利息残高推移表.csv")
        DoEvents
        Call MX040_未払利息残高推移表(GKeyName & "_" & "貸付未払利息残高推移表.csv")
    
    Case "平均金利平均残高推移表"
    
        kubun = "月次"
        KinNo = GRpt.金融
        
        Call MX040_平均金利平均残高推移表(GKeyName & "_" & "平均金利平均残高推移表.csv")
    
    Case "借入残高表"
    
        kubun = "月次"
        KinNo = GRpt.金融
        
        If GRpt.推移 = "年次" Then
            ws01 = GRpt.テキスト_01 & "_" & GRpt.推移 & "_"
        Else
            If G基本情報.日付入力区分 = "0" Then
            '和暦入力
                ws01 = GRpt.テキスト_02 & "_" & GRpt.推移 & "_"
            Else
            '西暦入力
                wsFD = Format(GRpt.テキスト_02, "yyyymm")
                ws01 = wsFD & "_" & GRpt.推移 & "_"
            End If
        End If
        
        Call MX040_借入残高表(GKeyName & "_" & ws01 & "借入残高表.csv")
    
    Case "社債残高表"
    
        kubun = "月次"
        KinNo = GRpt.金融
        
        If GRpt.推移 = "年次" Then
            ws01 = GRpt.テキスト_01 & "_" & GRpt.推移 & "_"
        Else
            If G基本情報.日付入力区分 = "0" Then
            '和暦入力
                ws01 = GRpt.テキスト_02 & "_" & GRpt.推移 & "_"
            Else
            '西暦入力
                wsFD = Format(GRpt.テキスト_02, "yyyymm")
                ws01 = wsFD & "_" & GRpt.推移 & "_"
            End If
        End If
        
        Call MX040_社債残高表(GKeyName & "_" & ws01 & "社債残高表.csv")
    
    Case "利息残高表"
    
        kubun = "月次"
        KinNo = GRpt.金融
        
        If GRpt.推移 = "年次" Then
            ws01 = GRpt.テキスト_01 & "_" & GRpt.推移 & "_"
        Else
            If G基本情報.日付入力区分 = "0" Then
            '和暦入力
                ws01 = GRpt.テキスト_02 & "_" & GRpt.推移 & "_"
            Else
            '西暦入力
                wsFD = Format(GRpt.テキスト_02, "yyyymm")
                ws01 = wsFD & "_" & GRpt.推移 & "_"
            End If
        End If
        
        Call MX040_前払利息残高表(GKeyName & "_" & ws01 & "利息前払残高表.csv")
        Call MX040_未払利息残高表(GKeyName & "_" & ws01 & "利息未払残高表.csv")
    
    Case "平均金利平均残高表"
    
        kubun = "月次"
        KinNo = GRpt.金融
        
        If GRpt.推移 = "年次" Then
            ws01 = GRpt.テキスト_01 & "_" & GRpt.推移 & "_"
        Else
            If G基本情報.日付入力区分 = "0" Then
            '和暦入力
                ws01 = GRpt.テキスト_02 & "_" & GRpt.推移 & "_"
            Else
            '西暦入力
                wsFD = Format(GRpt.テキスト_02, "yyyymm")
                ws01 = wsFD & "_" & GRpt.推移 & "_"
            End If
        End If
        
        Call MX040_平均金利平均残高表(GKeyName & "_" & ws01 & "平均金利平均残高表.csv")
    
    Case "金融機関別残高表"
        
        If GRpt.推移 = "年次" Then
            ws01 = GRpt.テキスト_01 & "_"
        Else
            If G基本情報.日付入力区分 = "0" Then
            '和暦入力
                ws01 = GRpt.テキスト_02 & "_"
            Else
            '西暦入力
                wsFD = Format(GRpt.テキスト_02, "yyyymm")
                ws01 = wsFD & "_"
            End If
        End If
        
        Call MX040_金融機関別残高表(GKeyName & "_" & ws01 & "金融機関別残高表.csv")
    
    Case "損益利息一覧表"
    
        kubun = "月次"
        KinNo = GRpt.金融
        
        If GRpt.推移 = "年次" Then
            ws01 = GRpt.テキスト_01 & "_" & GRpt.推移 & "_"
        Else
            If G基本情報.日付入力区分 = "0" Then
            '和暦入力
                ws01 = GRpt.テキスト_02 & "_" & GRpt.推移 & "_"
            Else
            '西暦入力
                wsFD = Format(GRpt.テキスト_02, "yyyymm")
                ws01 = wsFD & "_" & GRpt.推移 & "_"
            End If
        End If
        
        Call MX040_損益利息一覧表(GKeyName & "_" & ws01 & "損益利息一覧表.csv")
    
    Case Else
        Exit Sub
    End Select
'
    If FLG_Check = True Then
       'MsgBox "対象ファイル起動中の為、CSVファイルを作成できませんでした。", vbOKOnly + vbInformation
    End If
    
    FLG_Check = False
'    Screen.MousePointer = Default
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MX040_CsvOut_KARISUII_ERR:
    pERR_MES = pPROGRAM_ID + "/ MX040_CsvOut_KARISUII() でエラー" + vbCrLf + vbCrLf + _
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
' MX040_借入残高推移表
'------------------------------------------------
Private Sub MX040_借入残高推移表(pCsvFileName As String)
'
    Dim j As Integer, wML As Integer
    Dim ws01 As String, wsNendo As String
    Dim wsTbl As String
'
    On Error Resume Next
'
    'フィールド数
    Select Case GRpt.推移
    Case "月次", "四半期"
        wML = 12
    Case "年次", "半期"
        wML = 10
    End Select

    wsTbl = "DBDA010_借入金"
    If GRpt.帳票名 = "借入残高推移表" Then
        wsTbl = "DBDA010_借入金"
    ElseIf GRpt.帳票名 = "貸付残高推移表" Then
        wsTbl = "DBDA010_貸付金"
    End If
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wstr = "SELECT "
    
    If GRpt.帳票名 = "借入残高推移表" Then
        wstr = wstr & "Z.借入番号,"
    ElseIf GRpt.帳票名 = "貸付残高推移表" Then
        wstr = wstr & "Z.借入番号 As 貸付番号,"
    End If
        
    wstr = wstr & " K.借入内容,"
    wstr = wstr & " KS.借入金種別名,"
    wstr = wstr & " K.銀行番号,"
    wstr = wstr & " G.銀行名,"
    wstr = wstr + "B.部門名,"
    
    wstr = wstr & "IIF(K.営業日区分 = " & P8.FCDbl(XMXA020_区分("営業日", "翌営業日")) & ",'翌営業日','前営業日') As 営業日,"
    wstr = wstr & "IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As 利息区分,"
    wstr = wstr & "IIF(K.利息計算日数区分 = " & P8.FCDbl(XMXA020_区分("利息計算日数", "営業日数")) & ",'営業日数','固定日数') As 利息計算日数,"
    wstr = wstr & "IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As 金利種別,"
    wstr = wstr & "KK.基準金利名,"
    wstr = wstr & "K.金利条件 As 金利備考,"
    
    wstr = wstr & "IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期借入金','長期借入金') AS 長短区分,"
    wstr = wstr & "IIF(K.有担保フラグ = " & P8.FCDbl(XMXA020_区分("有担フラグ", "無担保")) & ",'無担保','有担保') As 担保区分,"
    wstr = wstr & "K.担保名 As 担保内容,"
    wstr = wstr & "IIF(設備フラグ=1,'設備','運転資金') As 設備区分,"
    wstr = wstr & "K.資金用途,"
    
    wstr = wstr & "KG.金利グループ名,"
    wstr = wstr & "IIF(K.sm区分=1 and K.金融リストラ番号<>'','借入SM',IIF(K.sm区分=0 and not (K.金融解約実行日 is null),'解約SM','')) As ｼﾐｭﾚｰｼｮﾝ内容,"
    wstr = wstr & "K.金融リストラ番号 As ｼﾐｭﾚｰｼｮﾝ番号,"
'    wstr = wstr & "format(K.金融解約実行日,'yyyy/mm/dd') As ｼﾐｭﾚｰｼｮﾝ解約日,"
    
    'wstr = wstr & "Format(K.実行日,'yyyy/mm/dd') As 実行日,"
    'wstr = wstr & "Format(K.初回返済年月,'yyyy/mm/dd') As 初回返済年月,"
    'wstr = wstr & "Format(K.初回返済実行日,'yyyy/mm/dd') As 初回返済年月日,"
    'wstr = wstr & "Format(K.最終返済年月,'yyyy/mm/dd') As 最終返済年月,"
    'wstr = wstr & "Format(K.最終返済実行日,'yyyy/mm/dd') As 最終返済年月日,"
    'wstr = wstr & "Format(K.解約実行日,'yyyy/mm/dd') As 解約年月日,"
    
    '2012.10.23　追加 by k.kunita
    wstr = wstr & "format(K.金融解約実行日,'" & Gfmtcsv年月日 & "') As ｼﾐｭﾚｰｼｮﾝ解約日,"
    
    wstr = wstr & "Format(K.実行日,'" & Gfmtcsv年月日 & "') As 実行日,"
    wstr = wstr & "Format(K.初回返済年月,'" & Gfmtcsv年月日 & "') As 初回返済年月,"
    wstr = wstr & "Format(K.初回返済実行日,'" & Gfmtcsv年月日 & "') As 初回返済年月日,"
    wstr = wstr & "Format(K.最終返済年月,'" & Gfmtcsv年月日 & "') As 最終返済年月,"
    wstr = wstr & "Format(K.最終返済実行日,'" & Gfmtcsv年月日 & "') As 最終返済年月日,"
    wstr = wstr & "Format(K.解約実行日,'" & Gfmtcsv年月日 & "') As 解約年月日,"
    
    wstr = wstr & "Format(IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.初回返済額,0),'#,##0') As 初回返済額,"
    wstr = wstr & "Format(IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.毎月返済額,0),'#,##0') As 毎月返済額,"
    wstr = wstr & "Format(IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.最終返済額,0),'#,##0') As 最終返済額,"
    
    'wstr = wstr + "IIF(K.変動利率フラグ=1,'*','') As 変動利率フラグ,"
    'wstr = wstr & "K.利率 As 利率,"
    wstr = wstr & "K.返済単位月数,"
    
    wstr = wstr & "Format(融資合計,'#,##0') As 合計融資金額,"
    wstr = wstr & "Format(元金合計,'#,##0') As 合計元金額,"
    wstr = wstr & "Format(利息合計,'#,##0') As 合計利息額,"
    wstr = wstr & "Format(返済合計,'#,##0') As 合計返済金額,"
    wstr = wstr & "Format(解約合計,'#,##0') As 合計解約金額,"
    wstr = wstr & "Format(残高合計,'#,##0') As 合計融資残高,"
    wstr = wstr & "Format(初期手数料合計+元金手数料合計+利息手数料合計,'#,##0') As 合計手数料,"
    wstr = wstr & "Format(保証合計,'#,##0') As 合計保証料,"
    'wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "', Format(前払利息減合計,'#,##0'), Format(未払利息増合計,'#,##0')) As 合計損益計算書計上利息額,"
    wstr = wstr & "Format(損益利息額合計,'#,##0') As 合計損益計算書計上利息額,"
    wstr = wstr & "Format(長短振替額合計,'#,##0') As 合計長短振替額,"
        
    For j = 1 To wML - 1
        ws01 = "_" & Right("0" & CStr(j), 2)
        
        If G基本情報.日付入力区分 = "0" Then
        '和暦入力
            wstr = wstr & "Format(融資" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "融資金額,"
            wstr = wstr & "Format(元金" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "元金額,"
            wstr = wstr & "Format(利息" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "利息額,"
            wstr = wstr & "Format(返済" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "返済金額,"
            wstr = wstr & "Format(解約" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "解約金額,"
            wstr = wstr & "Format(残高" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "融資残高,"
            wstr = wstr & "Format(初期手数料" & ws01 & "+元金手数料" & ws01 & "+利息手数料" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "手数料,"
            wstr = wstr & "Format(保証" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "保証料,"
            
            'wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Format(前払利息減" & ws01 & ",'#,##0'),Format(未払利息増" & ws01 & ",'#,##0')) " & " As " & w推移表タイトル.X番目年月(j) & "損益計算書計上利息額,"
            wstr = wstr & "Format(損益利息額" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "損益計算書計上利息額,"
            wstr = wstr & "Format(長短振替額" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "長短振替額,"
        Else
        '西暦入力
            wstr = wstr & "Format(融資" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "融資金額,"
            wstr = wstr & "Format(元金" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "元金額,"
            wstr = wstr & "Format(利息" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "利息額,"
            wstr = wstr & "Format(返済" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "返済金額,"
            wstr = wstr & "Format(解約" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "解約金額,"
            wstr = wstr & "Format(残高" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "融資残高,"
            wstr = wstr & "Format(初期手数料" & ws01 & "+元金手数料" & ws01 & "+利息手数料" & ws01 & ",'#,##0')  As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "手数料,"
            wstr = wstr & "Format(保証" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "保証料,"
            'wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "' , Format(前払利息減" & ws01 & ",'#,##0') , Format(未払利息増" & ws01 & ",'#,##0') ) As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "損益計算書計上利息額,"
            wstr = wstr & "Format(損益利息額" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "損益計算書計上利息額,"
            wstr = wstr & "Format(長短振替額" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "長短振替額,"
        
        End If
        
    Next j
    
    For j = wML To wML
        ws01 = "_" & Right("0" & CStr(j), 2)
        
        If G基本情報.日付入力区分 = "0" Then
        '和暦入力
            wstr = wstr & "Format(融資" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "融資金額,"
            wstr = wstr & "Format(元金" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "元金額,"
            wstr = wstr & "Format(利息" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "利息額,"
            wstr = wstr & "Format(返済" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "返済金額,"
            wstr = wstr & "Format(解約" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "解約金額,"
            wstr = wstr & "Format(残高" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "融資残高,"
            wstr = wstr & "Format(初期手数料" & ws01 & "+元金手数料" & ws01 & "+利息手数料" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "手数料,"
            wstr = wstr & "Format(保証" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "保証料,"
            
            'wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Format(前払利息減" & ws01 & ",'#,##0'),Format(未払利息増" & ws01 & ",'#,##0')) " & " As " & w推移表タイトル.X番目年月(j) & "損益計算書計上利息額,"
            wstr = wstr & "Format(損益利息額" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "損益計算書計上利息額,"
            wstr = wstr & "Format(長短振替額" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "長短振替額"
        Else
        '西暦入力
            wstr = wstr & "Format(融資" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "融資金額,"
            wstr = wstr & "Format(元金" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "元金額,"
            wstr = wstr & "Format(利息" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "利息額,"
            wstr = wstr & "Format(返済" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "返済金額,"
            wstr = wstr & "Format(解約" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "解約金額,"
            wstr = wstr & "Format(残高" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "融資残高,"
            wstr = wstr & "Format(初期手数料" & ws01 & "+元金手数料" & ws01 & "+利息手数料" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "手数料,"
            wstr = wstr & "Format(保証" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "保証料,"
        
            'wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "' , Format(前払利息減" & ws01 & ",'#,##0') , Format(未払利息増" & ws01 & ",'#,##0') ) As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "損益計算書計上利息額,"
            wstr = wstr & "Format(損益利息額" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "損益計算書計上利息額,"
            wstr = wstr & "Format(長短振替額" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "長短振替額"
        End If
        
    Next j
    
    wstr = wstr & " INTO [Text;database=" & wCsvDir & "]." & "[" & pCsvFileName & "]"
    
    wstr = wstr + " FROM ((((((DCDA010_借入残高推移表結果 As Z"
    wstr = wstr & " INNER JOIN DCDA010_借入残高推移表結果2 As Z2"
    wstr = wstr & " ON Z.借入番号=Z2.借入番号)"
    wstr = wstr + " Inner Join " & wsTbl & " As K"
    wstr = wstr + " ON Z.借入番号 = K.借入番号)"
    wstr = wstr + " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr + " ON K.銀行番号 = G.銀行番号)"
    wstr = wstr + " Inner Join DAAA116_借入金種別 As KS"
    wstr = wstr + " ON K.借入金種別区分 = KS.借入金種別区分)"
    wstr = wstr + " LEFT JOIN DAAA200_部門マスタ AS B"
    wstr = wstr + "  ON B.部門番号 = K.プロジェクト番号)"
    wstr = wstr + " Left Join DAAA116_基準金利 As KK"
    wstr = wstr + " ON K.基準金利区分 = KK.基準金利区分)"
    wstr = wstr + " Left Join DAAA115_金利シミュレーショングループ As KG"
    wstr = wstr + " ON K.金利グループ区分 = KG.金利グループ区分"
    
    'Order
    wstr = wstr & " Order By K.銀行番号,Z.借入番号"
    
    GDb.Execute wstr
    
    DoEvents
    
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
End Sub

'------------------------------------------------
' MX040_借入残高表
'------------------------------------------------
Private Sub MX040_借入残高表(pCsvFileName As String)
'
    Dim j As Integer
    Dim wd01 As Double
    Dim ws01 As String, wsNendo As String
    Dim wsTbl As String, wstr2 As String
    
    Dim wsF01 As String, wsF02 As String
'
    On Error Resume Next
'
    If GRpt.推移 = "月次" Then
         wsF02 = "当月融資金額"
    Else
         wsF02 = "当期融資金額"
    End If
    
    wsTbl = "DBDA010_借入金"
    If GRpt.帳票名 = "借入残高表" Then
        wsTbl = "DBDA010_借入金"
    
        wsF01 = "借入番号"
    ElseIf GRpt.帳票名 = "貸付残高表" Then
        wsTbl = "DBDA010_貸付金"
    
        wsF01 = "貸付番号"
    End If
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    'GInt1パラメータ
    ws01 = "_" & Right("00" & CStr(GInt1), 2)
    
    wstr = "SELECT "
    wstr = wstr & "Z.借入番号,"
    wstr = wstr & "K.借入内容,"
    wstr = wstr & "KS.借入金種別名,"
    
    wstr = wstr & "K.銀行番号,"
    wstr = wstr & "G.銀行名,"
    wstr = wstr + "B.部門名,"
    
    wstr = wstr & "IIF(K.営業日区分 = " & P8.FCDbl(XMXA020_区分("営業日", "翌営業日")) & ",'翌営業日','前営業日') As 営業日,"
    wstr = wstr & "IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As 利息区分,"
    wstr = wstr & "IIF(K.利息計算日数区分 = " & P8.FCDbl(XMXA020_区分("利息計算日数", "営業日数")) & ",'営業日数','固定日数') As 利息計算日数,"
    wstr = wstr & "IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As 金利種別,"
    wstr = wstr & "KK.基準金利名,"
    wstr = wstr & "K.金利条件 As 金利備考,"
    
    wstr = wstr & "IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期借入金','長期借入金') AS 長短区分,"
    wstr = wstr & "IIF(K.有担保フラグ = " & P8.FCDbl(XMXA020_区分("有担フラグ", "無担保")) & ",'無担保','有担保') As 担保区分,"
    wstr = wstr & "K.担保名 As 担保内容,"
    wstr = wstr & "IIF(設備フラグ=1,'設備','運転資金') As 設備区分,"
    wstr = wstr & "K.資金用途,"
    
    wstr = wstr & "KG.金利グループ名,"
    wstr = wstr & "IIF(K.sm区分=1 and K.金融リストラ番号<>'','借入SM',IIF(K.sm区分=0 and not (K.金融解約実行日 is null),'解約SM','')) As ｼﾐｭﾚｰｼｮﾝ内容,"
    wstr = wstr & "K.金融リストラ番号 As ｼﾐｭﾚｰｼｮﾝ番号,"
'    wstr = wstr & "format(K.金融解約実行日,'yyyy/mm/dd') As ｼﾐｭﾚｰｼｮﾝ解約日,"
    
'    wstr = wstr & "Format(K.実行日,'yyyy/mm/dd') As 実行日,"
'    wstr = wstr & "Format(K.初回返済年月,'yyyy/mm/dd') As 初回返済年月,"
'    wstr = wstr & "Format(K.初回返済実行日,'yyyy/mm/dd') As 初回返済年月日,"
'    wstr = wstr & "Format(K.最終返済年月,'yyyy/mm/dd') As 最終返済年月,"
'    wstr = wstr & "Format(K.最終返済実行日,'yyyy/mm/dd') As 最終返済年月日,"
'    wstr = wstr & "Format(K.解約実行日,'yyyy/mm/dd') As 解約年月日,"
    
    '2012.10.24　追加 by k.kuni
    wstr = wstr & "format(K.金融解約実行日,'" & Gfmtcsv年月日 & "') As ｼﾐｭﾚｰｼｮﾝ解約日,"
    
    wstr = wstr & "Format(K.実行日,'" & Gfmtcsv年月日 & "') As 実行日,"
    wstr = wstr & "Format(K.初回返済年月,'" & Gfmtcsv年月日 & "') As 初回返済年月,"
    wstr = wstr & "Format(K.初回返済実行日,'" & Gfmtcsv年月日 & "') As 初回返済年月日,"
    wstr = wstr & "Format(K.最終返済年月,'" & Gfmtcsv年月日 & "') As 最終返済年月,"
    wstr = wstr & "Format(K.最終返済実行日,'" & Gfmtcsv年月日 & "') As 最終返済年月日,"
    wstr = wstr & "Format(K.解約実行日,'" & Gfmtcsv年月日 & "') As 解約年月日,"
    
    wstr = wstr & "IIF(KS.利子補給金フラグ=0,IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.初回返済額,0),0) As 初回返済額,"
    wstr = wstr & "IIF(KS.利子補給金フラグ=0,IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.毎月返済額,0),0) As 毎月返済額,"
    wstr = wstr & "IIF(KS.利子補給金フラグ=0,IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.最終返済額,0),0) As 最終返済額,"
    
    'wstr = wstr + "IIF(K.変動利率フラグ=1,'*','') As 変動利率フラグ,"
    'wstr = wstr & "K.利率 As 利率,"
    wstr = wstr & "K.返済単位月数,"
    wstr = wstr & "format(利率" & ws01 & ",'#,##0.00000') As 利率,"
    
    '利子補給に伴う変更
    'wstr = wstr & "K.融資金額 As 融資金額,"
    wstr = wstr & "IIF(KS.利子補給金フラグ=0,K.融資金額,0) As 融資金額,"
    
    wstr = wstr & "融資" & ws01 & " As 当月融資金額,"
    wstr = wstr & "残高" & ws01 & " As 融資残高,"
    wstr = wstr & "元金" & ws01 & " As 支払元金額,"
    wstr = wstr & "利息" & ws01 & " As 支払利息額,"
    wstr = wstr & "返済" & ws01 & " As 支払額,"
    wstr = wstr & "解約" & ws01 & " As 解約金額,"
    'wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',前払利息減" & ws01 & ",未払利息増" & ws01 & ") " & " As 損益利息,"
    wstr = wstr & "長短振替額" & ws01 & " As 長短振替,"
    wstr = wstr & "損益利息額" & ws01 & " As 損益利息"
    
    wstr = wstr + " FROM ((((((DCDA010_借入残高推移表結果 As Z"
    wstr = wstr & " INNER JOIN DCDA010_借入残高推移表結果2 As Z2"
    wstr = wstr & " ON Z.借入番号=Z2.借入番号)"
    wstr = wstr + " Inner Join " & wsTbl & " As K"
    wstr = wstr + " ON Z.借入番号 = K.借入番号)"
    wstr = wstr + " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr + " ON K.銀行番号 = G.銀行番号)"
    wstr = wstr + " Inner Join DAAA116_借入金種別 As KS"
    wstr = wstr + " ON K.借入金種別区分 = KS.借入金種別区分)"
    wstr = wstr + " LEFT JOIN DAAA200_部門マスタ AS B"
    wstr = wstr + "  ON B.部門番号 = K.プロジェクト番号)"
    wstr = wstr + " Left Join DAAA116_基準金利 As KK"
    wstr = wstr + " ON K.基準金利区分 = KK.基準金利区分)"
    wstr = wstr + " Left Join DAAA115_金利シミュレーショングループ As KG"
    wstr = wstr + " ON K.金利グループ区分 = KG.金利グループ区分"
    
    wstr = wstr & " Where 融資" + ws01 + "<>0"
    wstr = wstr & " Or 元金" + ws01 + "<>0"
    wstr = wstr & " Or 利息" + ws01 + "<>0"
    wstr = wstr & " Or 返済" + ws01 + "<>0"
    wstr = wstr & " Or 解約" + ws01 + "<>0"
    wstr = wstr & " Or 残高" + ws01 + "<>0"
    wstr = wstr & " Or 長短振替額" + ws01 + "<>0"
    wstr = wstr & " Or 損益利息額" + ws01 + "<>0"
    
    wstr = wstr & " Order By K.借入金種別区分,K.銀行番号,Z.借入番号"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        
        '名称
        Write #1, _
            "借入番号", "借入内容", "借入金種別名", "銀行番号", "銀行名", "部門名", _
            "営業日", "利息区分", "利息計算日数", "金利種別", "基準金利名", "金利備考", _
            "長短区分", "担保区分", "担保内容", "設備区分", "資金用途", _
            "金利グループ名", "ｼﾐｭﾚｰｼｮﾝ内容", "ｼﾐｭﾚｰｼｮﾝ番号", "ｼﾐｭﾚｰｼｮﾝ解約日", _
            "実行日", "初回返済年月", "初回返済年月日", "最終返済年月", "最終返済年月日", "解約年月日", _
            "初回返済額", "毎月返済額", "最終返済額", _
            "返済単位月数", "利率", _
            "融資金額", _
            "融資残高", _
            "支払元金額", _
            "支払利息額", _
            "支払額", _
            "長短振替額", _
            "損益計算書計上利息額", _
            "返済率"
    
    Do Until wRs.eof
    
        wd01 = Format(Round(P8.FCDiv(P8.FCDbl(wRs.Fields("融資金額").Value) - P8.FCDbl(wRs.Fields("融資残高").Value), P8.FCDbl(wRs.Fields("融資金額").Value)) * 100, 3), "#,##0.00")
        
        Write #1, _
            P8.FCStr(wRs.Fields("借入番号").Value), P8.FCStr(wRs.Fields("借入内容").Value), _
            P8.FCStr(wRs.Fields("借入金種別名").Value), _
            P8.FCStr(wRs.Fields("銀行番号").Value), P8.FCStr(wRs.Fields("銀行名").Value), P8.FCStr(wRs.Fields("部門名").Value), _
            P8.FCStr(wRs.Fields("営業日").Value), P8.FCStr(wRs.Fields("利息区分").Value), P8.FCStr(wRs.Fields("利息計算日数").Value), _
            P8.FCStr(wRs.Fields("金利種別").Value), P8.FCStr(wRs.Fields("基準金利名").Value), P8.FCStr(wRs.Fields("金利備考").Value), _
            P8.FCStr(wRs.Fields("長短区分").Value), P8.FCStr(wRs.Fields("担保区分").Value), P8.FCStr(wRs.Fields("担保内容").Value), _
            P8.FCStr(wRs.Fields("設備区分").Value), P8.FCStr(wRs.Fields("資金用途").Value), _
            P8.FCStr(wRs.Fields("金利グループ名").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ内容").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ番号").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ解約日").Value), _
            P8.FCStr(wRs.Fields("実行日").Value), _
            P8.FCStr(wRs.Fields("初回返済年月").Value), P8.FCStr(wRs.Fields("初回返済年月日").Value), _
            P8.FCStr(wRs.Fields("最終返済年月").Value), P8.FCStr(wRs.Fields("最終返済年月日").Value), _
            P8.FCStr(wRs.Fields("解約年月日").Value), _
            P8.FCDbl(wRs.Fields("初回返済額").Value), P8.FCDbl(wRs.Fields("毎月返済額").Value), P8.FCDbl(wRs.Fields("最終返済額").Value), _
            P8.FCDbl(wRs.Fields("返済単位月数").Value), P8.FCDbl(wRs.Fields("利率").Value), _
            P8.FCDbl(wRs.Fields("融資金額").Value), _
            P8.FCDbl(wRs.Fields("融資残高").Value), _
            P8.FCDbl(wRs.Fields("支払元金額").Value), _
            P8.FCDbl(wRs.Fields("支払利息額").Value), _
            P8.FCDbl(wRs.Fields("支払額").Value), _
            P8.FCDbl(wRs.Fields("長短振替").Value), _
            P8.FCDbl(wRs.Fields("損益利息").Value), _
            wd01

    wRs.MoveNext
    Loop
    End If
    wRs.Close
    Set wRs = Nothing
    
    '合計
    wstr2 = "select "
    wstr2 = wstr2 & "'総合計' As タイトル,'',count(Z.借入番号) & '件' As 件数,'','',"
    wstr2 = wstr2 & "'','','','','','','','','','','','','','','','','','','','','','','','','','','',"

    wstr2 = wstr2 & "Sum(K.融資金額) As 融資金額,"
    wstr2 = wstr2 & "Sum(融資" & ws01 & ") As 当月融資金額,"
    wstr2 = wstr2 & "Sum(残高" & ws01 & ") As 融資残高,"
    wstr2 = wstr2 & "Sum(元金" & ws01 & ") As 支払元金額,"
    wstr2 = wstr2 & "Sum(利息" & ws01 & ") As 支払利息額,"
    wstr2 = wstr2 & "Sum(返済" & ws01 & ") As 支払額,"
    wstr2 = wstr2 & "Sum(解約" & ws01 & ") As 解約金額,"
    'wstr2 = wstr2 & "Sum(IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',前払利息減" & ws01 & ",未払利息増" & ws01 & ")) As 損益利息,"
    wstr2 = wstr2 & "Sum(損益利息額" & ws01 & ") As 損益利息,"
    wstr2 = wstr2 & "Sum(長短振替額" & ws01 & ") As 長短振替"
    
    wstr2 = wstr2 + " FROM (DCDA010_借入残高推移表結果 As Z"
    wstr2 = wstr2 + " Inner Join DCDA010_借入残高推移表結果2 As Z2"
    wstr2 = wstr2 + " ON Z.借入番号 = Z2.借入番号)"
    wstr2 = wstr2 + " Inner Join " & wsTbl & " As K"
    wstr2 = wstr2 + " ON Z.借入番号 = K.借入番号"

    wstr2 = wstr2 & " Where 融資" + ws01 + "<>0"
    wstr2 = wstr2 & " Or 元金" + ws01 + "<>0"
    wstr2 = wstr2 & " Or 利息" + ws01 + "<>0"
    wstr2 = wstr2 & " Or 返済" + ws01 + "<>0"
    wstr2 = wstr2 & " Or 解約" + ws01 + "<>0"
    wstr2 = wstr2 & " Or 残高" + ws01 + "<>0"
    wstr2 = wstr2 & " Or 損益利息額" + ws01 + "<>0"
    wstr2 = wstr2 & " Or 長短振替額" + ws01 + "<>0"
    Call AdoRecordsetOpen(GDb, wRs, wstr2)
    If Not wRs.eof Then
    Do Until wRs.eof
    
        wd01 = Format(Round(P8.FCDiv(P8.FCDbl(wRs.Fields("融資金額").Value) - P8.FCDbl(wRs.Fields("融資残高").Value), P8.FCDbl(wRs.Fields("融資金額").Value)) * 100, 3), "#,##0.00")
        
        Write #1, _
            P8.FCStr(wRs.Fields("タイトル").Value), "", P8.FCStr(wRs.Fields("件数").Value), _
            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
            P8.FCDbl(wRs.Fields("融資金額").Value), _
            P8.FCDbl(wRs.Fields("融資残高").Value), _
            P8.FCDbl(wRs.Fields("支払元金額").Value), _
            P8.FCDbl(wRs.Fields("支払利息額").Value), _
            P8.FCDbl(wRs.Fields("支払額").Value), _
            P8.FCDbl(wRs.Fields("長短振替").Value), _
            P8.FCDbl(wRs.Fields("損益利息").Value), _
            wd01

    wRs.MoveNext
    Loop
    End If
    wRs.Close
    Set wRs = Nothing
        
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_社債残高推移表
'------------------------------------------------
Private Sub MX040_社債残高推移表(pCsvFileName As String)
'
    Dim j As Integer, wML As Integer
    Dim ws01 As String, wsNendo As String
    Dim wsTbl As String
'
    On Error Resume Next
'
    'フィールド数
    Select Case GRpt.推移
    Case "月次", "四半期"
        wML = 12
    Case "年次", "半期"
        wML = 10
    End Select

    wsTbl = "DBDA010_借入金"
    If GRpt.帳票名 = "借入残高推移表" Then
        wsTbl = "DBDA010_借入金"
    ElseIf GRpt.帳票名 = "貸付残高推移表" Then
        wsTbl = "DBDA010_貸付金"
    ElseIf GRpt.帳票名 = "社債残高推移表" Then
        wsTbl = "DBDA010_借入金"
    End If
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wstr = "SELECT "
    
    If GRpt.帳票名 = "借入残高推移表" Then
        wstr = wstr & "Z.借入番号,"
    ElseIf GRpt.帳票名 = "社債残高推移表" Then
        wstr = wstr & "Z.借入番号,"
    ElseIf GRpt.帳票名 = "貸付残高推移表" Then
        wstr = wstr & "Z.借入番号 As 貸付番号,"
    End If
        
    wstr = wstr & " K.借入内容,"
    wstr = wstr & " KS.借入金種別名,"
    wstr = wstr & " K.銀行番号,"
    wstr = wstr & " G.銀行名,"
    wstr = wstr + "B.部門名,"
    
    wstr = wstr & "IIF(K.営業日区分 = " & P8.FCDbl(XMXA020_区分("営業日", "翌営業日")) & ",'翌営業日','前営業日') As 営業日,"
    wstr = wstr & "IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As 利息区分,"
    wstr = wstr & "IIF(K.利息計算日数区分 = " & P8.FCDbl(XMXA020_区分("利息計算日数", "営業日数")) & ",'営業日数','固定日数') As 利息計算日数,"
    wstr = wstr & "IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As 金利種別,"
    wstr = wstr & "KK.基準金利名,"
    wstr = wstr & "K.金利条件 As 金利備考,"
    
    wstr = wstr & "IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期借入金','長期借入金') AS 長短区分,"
    wstr = wstr & "IIF(K.有担保フラグ = " & P8.FCDbl(XMXA020_区分("有担フラグ", "無担保")) & ",'無担保','有担保') As 担保区分,"
    wstr = wstr & "K.担保名 As 担保内容,"
    wstr = wstr & "IIF(設備フラグ=1,'設備','運転資金') As 設備区分,"
    wstr = wstr & "K.資金用途,"
    
    wstr = wstr & "KG.金利グループ名,"
    wstr = wstr & "IIF(K.sm区分=1 and K.金融リストラ番号<>'','借入SM',IIF(K.sm区分=0 and not (K.金融解約実行日 is null),'解約SM','')) As ｼﾐｭﾚｰｼｮﾝ内容,"
    wstr = wstr & "K.金融リストラ番号 As ｼﾐｭﾚｰｼｮﾝ番号,"
    wstr = wstr & "format(K.金融解約実行日,'yyyy/mm/dd') As ｼﾐｭﾚｰｼｮﾝ解約日,"
    
    wstr = wstr & "Format(K.実行日,'yyyy/mm/dd') As 実行日,"
    wstr = wstr & "Format(K.初回返済年月,'yyyy/mm/dd') As 初回返済年月,"
    wstr = wstr & "Format(K.初回返済実行日,'yyyy/mm/dd') As 初回返済年月日,"
    wstr = wstr & "Format(K.最終返済年月,'yyyy/mm/dd') As 最終返済年月,"
    wstr = wstr & "Format(K.最終返済実行日,'yyyy/mm/dd') As 最終返済年月日,"
    wstr = wstr & "Format(K.解約実行日,'yyyy/mm/dd') As 解約年月日,"
    
    wstr = wstr & "Format(IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.初回返済額,0),'#,##0') As 初回返済額,"
    wstr = wstr & "Format(IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.毎月返済額,0),'#,##0') As 毎月返済額,"
    wstr = wstr & "Format(IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.最終返済額,0),'#,##0') As 最終返済額,"
    
    'wstr = wstr + "IIF(K.変動利率フラグ=1,'*','') As 変動利率フラグ,"
    'wstr = wstr & "K.利率 As 利率,"
    wstr = wstr & "K.返済単位月数,"
    
    wstr = wstr & "Format(融資合計,'#,##0') As 合計融資金額,"
    wstr = wstr & "Format(元金合計,'#,##0') As 合計元金額,"
    wstr = wstr & "Format(利息合計,'#,##0') As 合計利息額,"
    wstr = wstr & "Format(返済合計,'#,##0') As 合計返済金額,"
    wstr = wstr & "Format(解約合計,'#,##0') As 合計解約金額,"
    wstr = wstr & "Format(残高合計,'#,##0') As 合計融資残高,"
    wstr = wstr & "Format(初期手数料合計+元金手数料合計+利息手数料合計,'#,##0') As 合計手数料,"
    wstr = wstr & "Format(保証合計,'#,##0') As 合計保証料,"
        
    For j = 1 To wML - 1
        ws01 = "_" & Right("0" & CStr(j), 2)
        
        If G基本情報.日付入力区分 = "0" Then
        '和暦入力
            wstr = wstr & "Format(融資" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "融資金額,"
            wstr = wstr & "Format(元金" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "元金額,"
            wstr = wstr & "Format(利息" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "利息額,"
            wstr = wstr & "Format(返済" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "返済金額,"
            wstr = wstr & "Format(解約" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "解約金額,"
            wstr = wstr & "Format(残高" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "融資残高,"
            wstr = wstr & "Format(初期手数料" & ws01 & "+元金手数料" & ws01 & "+利息手数料" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "手数料,"
            wstr = wstr & "Format(保証" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "保証料,"
        Else
        '西暦入力
            wstr = wstr & "Format(融資" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "融資金額,"
            wstr = wstr & "Format(元金" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "元金額,"
            wstr = wstr & "Format(利息" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "利息額,"
            wstr = wstr & "Format(返済" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "返済金額,"
            wstr = wstr & "Format(解約" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "解約金額,"
            wstr = wstr & "Format(残高" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "融資残高,"
            wstr = wstr & "Format(初期手数料" & ws01 & "+元金手数料" & ws01 & "+利息手数料" & ws01 & ",'#,##0')  As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "手数料,"
            wstr = wstr & "Format(保証" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "保証料,"
        End If
        
    Next j
    
    For j = wML To wML
        ws01 = "_" & Right("0" & CStr(j), 2)
        
        If G基本情報.日付入力区分 = "0" Then
        '和暦入力
            wstr = wstr & "Format(融資" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "融資金額,"
            wstr = wstr & "Format(元金" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "元金額,"
            wstr = wstr & "Format(利息" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "利息額,"
            wstr = wstr & "Format(返済" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "返済金額,"
            wstr = wstr & "Format(解約" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "解約金額,"
            wstr = wstr & "Format(残高" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "融資残高,"
            wstr = wstr & "Format(初期手数料" & ws01 & "+元金手数料" & ws01 & "+利息手数料" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "手数料,"
            wstr = wstr & "Format(保証" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "保証料"
        Else
        '西暦入力
            wstr = wstr & "Format(融資" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "融資金額,"
            wstr = wstr & "Format(元金" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "元金額,"
            wstr = wstr & "Format(利息" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "利息額,"
            wstr = wstr & "Format(返済" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "返済金額,"
            wstr = wstr & "Format(解約" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "解約金額,"
            wstr = wstr & "Format(残高" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "融資残高,"
            wstr = wstr & "Format(初期手数料" & ws01 & "+元金手数料" & ws01 & "+利息手数料" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "手数料,"
            wstr = wstr & "Format(保証" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "保証料"
        End If
        
    Next j
    
    wstr = wstr & " INTO [Text;database=" & wCsvDir & "]." & "[" & pCsvFileName & "]"
    
    wstr = wstr + " FROM (((((DCDA010_借入残高推移表結果 As Z"
    wstr = wstr + " Inner Join " & wsTbl & " As K"
    wstr = wstr + " ON Z.借入番号 = K.借入番号)"
    wstr = wstr + " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr + " ON K.銀行番号 = G.銀行番号)"
    wstr = wstr + " Inner Join DAAA116_借入金種別 As KS"
    wstr = wstr + " ON K.借入金種別区分 = KS.借入金種別区分)"
    wstr = wstr + " LEFT JOIN DAAA200_部門マスタ AS B"
    wstr = wstr + "  ON B.部門番号 = K.プロジェクト番号)"
    wstr = wstr + " Left Join DAAA116_基準金利 As KK"
    wstr = wstr + " ON K.基準金利区分 = KK.基準金利区分)"
    wstr = wstr + " Left Join DAAA115_金利シミュレーショングループ As KG"
    wstr = wstr + " ON K.金利グループ区分 = KG.金利グループ区分"
    
    'Order
    wstr = wstr & " Order By K.銀行番号,Z.借入番号"
    
    GDb.Execute wstr
    
    DoEvents
    
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
End Sub

'------------------------------------------------
' MX040_社債残高表
'------------------------------------------------
Private Sub MX040_社債残高表(pCsvFileName As String)
'
    Dim j As Integer
    Dim wd01 As Double
    Dim ws01 As String, wsNendo As String
    Dim wsTbl As String, wstr2 As String
    
    Dim wsF01 As String, wsF02 As String
'
    On Error Resume Next
'
    If GRpt.推移 = "月次" Then
         wsF02 = "当月融資金額"
    Else
         wsF02 = "当期融資金額"
    End If
    
    wsTbl = "DBDA010_借入金"
    If GRpt.帳票名 = "借入残高表" Then
        wsTbl = "DBDA010_借入金"
    
        wsF01 = "借入番号"
    ElseIf GRpt.帳票名 = "社債残高表" Then
        wsTbl = "DBDA010_借入金"
    
        wsF01 = "借入番号"
    ElseIf GRpt.帳票名 = "貸付残高表" Then
        wsTbl = "DBDA010_貸付金"
    
        wsF01 = "貸付番号"
    End If
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    'GInt1パラメータ
    ws01 = "_" & Right("00" & CStr(GInt1), 2)
    
    wstr = "SELECT "
    wstr = wstr & "Z.借入番号,"
    wstr = wstr & "K.借入内容,"
    wstr = wstr & "KS.借入金種別名,"
    
    wstr = wstr & "K.銀行番号,"
    wstr = wstr & "G.銀行名,"
    wstr = wstr + "B.部門名,"
    
    wstr = wstr & "IIF(K.営業日区分 = " & P8.FCDbl(XMXA020_区分("営業日", "翌営業日")) & ",'翌営業日','前営業日') As 営業日,"
    wstr = wstr & "IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As 利息区分,"
    wstr = wstr & "IIF(K.利息計算日数区分 = " & P8.FCDbl(XMXA020_区分("利息計算日数", "営業日数")) & ",'営業日数','固定日数') As 利息計算日数,"
    wstr = wstr & "IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As 金利種別,"
    wstr = wstr & "KK.基準金利名,"
    wstr = wstr & "K.金利条件 As 金利備考,"
    
    wstr = wstr & "IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期借入金','長期借入金') AS 長短区分,"
    wstr = wstr & "IIF(K.有担保フラグ = " & P8.FCDbl(XMXA020_区分("有担フラグ", "無担保")) & ",'無担保','有担保') As 担保区分,"
    wstr = wstr & "K.担保名 As 担保内容,"
    wstr = wstr & "IIF(設備フラグ=1,'設備','運転資金') As 設備区分,"
    wstr = wstr & "K.資金用途,"
    
    wstr = wstr & "KG.金利グループ名,"
    wstr = wstr & "IIF(K.sm区分=1 and K.金融リストラ番号<>'','借入SM',IIF(K.sm区分=0 and not (K.金融解約実行日 is null),'解約SM','')) As ｼﾐｭﾚｰｼｮﾝ内容,"
    wstr = wstr & "K.金融リストラ番号 As ｼﾐｭﾚｰｼｮﾝ番号,"
    wstr = wstr & "format(K.金融解約実行日,'yyyy/mm/dd') As ｼﾐｭﾚｰｼｮﾝ解約日,"
    
    wstr = wstr & "Format(K.実行日,'yyyy/mm/dd') As 実行日,"
    wstr = wstr & "Format(K.初回返済年月,'yyyy/mm/dd') As 初回返済年月,"
    wstr = wstr & "Format(K.初回返済実行日,'yyyy/mm/dd') As 初回返済年月日,"
    wstr = wstr & "Format(K.最終返済年月,'yyyy/mm/dd') As 最終返済年月,"
    wstr = wstr & "Format(K.最終返済実行日,'yyyy/mm/dd') As 最終返済年月日,"
    wstr = wstr & "Format(K.解約実行日,'yyyy/mm/dd') As 解約年月日,"
    
    wstr = wstr & "IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.初回返済額,0) As 初回返済額,"
    wstr = wstr & "IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.毎月返済額,0) As 毎月返済額,"
    wstr = wstr & "IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.最終返済額,0) As 最終返済額,"
    
    'wstr = wstr + "IIF(K.変動利率フラグ=1,'*','') As 変動利率フラグ,"
    'wstr = wstr & "K.利率 As 利率,"
    wstr = wstr & "K.返済単位月数,"
    wstr = wstr & "format(利率" & ws01 & ",'#,##0.00000') As 利率,"
    
    wstr = wstr & "K.融資金額 As 融資金額,"
    wstr = wstr & "融資" & ws01 & " As 当月融資金額,"
    wstr = wstr & "元金" & ws01 & " As 支払元金額,"
    wstr = wstr & "利息" & ws01 & " As 支払利息額,"
    wstr = wstr & "返済" & ws01 & " As 支払額,"
    wstr = wstr & "解約" & ws01 & " As 解約金額,"
    wstr = wstr & "残高" & ws01 & " As 融資残高,"
    wstr = wstr & "初期手数料" & ws01 & "+元金手数料" & ws01 & "+利息手数料" & ws01 & " As 手数料,"
    wstr = wstr & "保証" & ws01 & " As 保証料"
    
    wstr = wstr + " FROM (((((DCDA010_借入残高推移表結果 As Z"
    wstr = wstr + " Inner Join " & wsTbl & " As K"
    wstr = wstr + " ON Z.借入番号 = K.借入番号)"
    wstr = wstr + " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr + " ON K.銀行番号 = G.銀行番号)"
    wstr = wstr + " Inner Join DAAA116_借入金種別 As KS"
    wstr = wstr + " ON K.借入金種別区分 = KS.借入金種別区分)"
    wstr = wstr + " LEFT JOIN DAAA200_部門マスタ AS B"
    wstr = wstr + "  ON B.部門番号 = K.プロジェクト番号)"
    wstr = wstr + " Left Join DAAA116_基準金利 As KK"
    wstr = wstr + " ON K.基準金利区分 = KK.基準金利区分)"
    wstr = wstr + " Left Join DAAA115_金利シミュレーショングループ As KG"
    wstr = wstr + " ON K.金利グループ区分 = KG.金利グループ区分"
    
    wstr = wstr & " Where 融資" + ws01 + "<>0"
    wstr = wstr & " Or 元金" + ws01 + "<>0"
    wstr = wstr & " Or 利息" + ws01 + "<>0"
    wstr = wstr & " Or 返済" + ws01 + "<>0"
    wstr = wstr & " Or 解約" + ws01 + "<>0"
    wstr = wstr & " Or 残高" + ws01 + "<>0"
    wstr = wstr & " Or 初期手数料" + ws01 + "<>0"
    wstr = wstr & " Or 元金手数料" + ws01 + "<>0"
    wstr = wstr & " Or 利息手数料" + ws01 + "<>0"
    wstr = wstr & " Or 保証" + ws01 + "<>0"
    
    wstr = wstr & " Order By K.銀行番号,Z.借入番号"
    
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        '名称
        Write #1, _
            "借入番号", "借入内容", "借入金種別名", "銀行番号", "銀行名", "部門名", _
            "営業日", "利息区分", "利息計算日数", "金利種別", "基準金利名", "金利備考", _
            "長短区分", "担保区分", "担保内容", "設備区分", "資金用途", _
            "金利グループ名", "ｼﾐｭﾚｰｼｮﾝ内容", "ｼﾐｭﾚｰｼｮﾝ番号", "ｼﾐｭﾚｰｼｮﾝ解約日", _
            "実行日", "初回返済年月", "初回返済年月日", "最終返済年月", "最終返済年月日", "解約年月日", _
            "初回返済額", "毎月返済額", "最終返済額", _
            "返済単位月数", "利率", _
            "融資金額", _
            "融資残高", _
            "支払元金額", _
            "支払利息額", _
            "支払額", _
            "返済率", _
            "手数料", "保証料"
    
    Do Until wRs.eof
    
        wd01 = Format(Round(P8.FCDiv(P8.FCDbl(wRs.Fields("融資金額").Value) - P8.FCDbl(wRs.Fields("融資残高").Value), P8.FCDbl(wRs.Fields("融資金額").Value)) * 100, 3), "#,##0.00")
        
        Write #1, _
            P8.FCStr(wRs.Fields("借入番号").Value), P8.FCStr(wRs.Fields("借入内容").Value), _
            P8.FCStr(wRs.Fields("借入金種別名").Value), _
            P8.FCStr(wRs.Fields("銀行番号").Value), P8.FCStr(wRs.Fields("銀行名").Value), P8.FCStr(wRs.Fields("部門名").Value), _
            P8.FCStr(wRs.Fields("営業日").Value), P8.FCStr(wRs.Fields("利息区分").Value), P8.FCStr(wRs.Fields("利息計算日数").Value), _
            P8.FCStr(wRs.Fields("金利種別").Value), P8.FCStr(wRs.Fields("基準金利名").Value), P8.FCStr(wRs.Fields("金利備考").Value), _
            P8.FCStr(wRs.Fields("長短区分").Value), P8.FCStr(wRs.Fields("担保区分").Value), P8.FCStr(wRs.Fields("担保内容").Value), _
            P8.FCStr(wRs.Fields("設備区分").Value), P8.FCStr(wRs.Fields("資金用途").Value), _
            P8.FCStr(wRs.Fields("金利グループ名").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ内容").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ番号").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ解約日").Value), _
            P8.FCStr(wRs.Fields("実行日").Value), _
            P8.FCStr(wRs.Fields("初回返済年月").Value), P8.FCStr(wRs.Fields("初回返済年月日").Value), _
            P8.FCStr(wRs.Fields("最終返済年月").Value), P8.FCStr(wRs.Fields("最終返済年月日").Value), _
            P8.FCStr(wRs.Fields("解約年月日").Value), _
            P8.FCDbl(wRs.Fields("初回返済額").Value), P8.FCDbl(wRs.Fields("毎月返済額").Value), P8.FCDbl(wRs.Fields("最終返済額").Value), _
            P8.FCDbl(wRs.Fields("返済単位月数").Value), P8.FCDbl(wRs.Fields("利率").Value), _
            P8.FCDbl(wRs.Fields("融資金額").Value), _
            P8.FCDbl(wRs.Fields("融資残高").Value), _
            P8.FCDbl(wRs.Fields("支払元金額").Value), _
            P8.FCDbl(wRs.Fields("支払利息額").Value), _
            P8.FCDbl(wRs.Fields("支払額").Value), _
            wd01, _
            P8.FCDbl(wRs.Fields("手数料").Value), _
            P8.FCDbl(wRs.Fields("保証料").Value)

    wRs.MoveNext
    Loop
    End If
    wRs.Close
    Set wRs = Nothing
    
    '合計
    wstr2 = "select "
    wstr2 = wstr2 & "'総合計' As タイトル,'',count(Z.借入番号) & '件' As 件数,'','',"
    wstr2 = wstr2 & "'','','','','','','','','','','','','','','','','','','','','','','','','','','',"
    
    wstr2 = wstr2 & "Sum(K.融資金額) As 融資金額,"
    wstr2 = wstr2 & "Sum(融資" & ws01 & ") As 当月融資金額,"
    wstr2 = wstr2 & "Sum(元金" & ws01 & ") As 支払元金額,"
    wstr2 = wstr2 & "Sum(利息" & ws01 & ") As 支払利息額,"
    wstr2 = wstr2 & "Sum(返済" & ws01 & ") As 支払額,"
    wstr2 = wstr2 & "Sum(解約" & ws01 & ") As 解約金額,"
    wstr2 = wstr2 & "Sum(残高" & ws01 & ") As 融資残高,"
    wstr2 = wstr2 & "Sum(初期手数料" & ws01 & "+元金手数料" & ws01 & "+利息手数料" & ws01 & ") As 手数料,"
    wstr2 = wstr2 & "Sum(保証" & ws01 & ") As 保証料"
    
    wstr2 = wstr2 + " FROM DCDA010_借入残高推移表結果 As Z"
    wstr2 = wstr2 + " Inner Join " & wsTbl & " As K"
    wstr2 = wstr2 + " ON Z.借入番号 = K.借入番号"
    
    wstr2 = wstr2 & " Where 融資" + ws01 + "<>0"
    wstr2 = wstr2 & " Or 元金" + ws01 + "<>0"
    wstr2 = wstr2 & " Or 利息" + ws01 + "<>0"
    wstr2 = wstr2 & " Or 返済" + ws01 + "<>0"
    wstr2 = wstr2 & " Or 解約" + ws01 + "<>0"
    wstr2 = wstr2 & " Or 残高" + ws01 + "<>0"
    wstr2 = wstr2 & " Or 初期手数料" + ws01 + "<>0"
    wstr2 = wstr2 & " Or 元金手数料" + ws01 + "<>0"
    wstr2 = wstr2 & " Or 利息手数料" + ws01 + "<>0"
    wstr2 = wstr2 & " Or 保証" + ws01 + "<>0"
    Call AdoRecordsetOpen(GDb, wRs, wstr2)
    If Not wRs.eof Then
    Do Until wRs.eof
    
        wd01 = Format(Round(P8.FCDiv(P8.FCDbl(wRs.Fields("融資金額").Value) - P8.FCDbl(wRs.Fields("融資残高").Value), P8.FCDbl(wRs.Fields("融資金額").Value)) * 100, 3), "#,##0.00")
        
        Write #1, _
            P8.FCStr(wRs.Fields("タイトル").Value), "", P8.FCStr(wRs.Fields("件数").Value), "", "", _
            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
            P8.FCDbl(wRs.Fields("融資金額").Value), _
            P8.FCDbl(wRs.Fields("融資残高").Value), _
            P8.FCDbl(wRs.Fields("支払元金額").Value), _
            P8.FCDbl(wRs.Fields("支払利息額").Value), _
            P8.FCDbl(wRs.Fields("支払額").Value), _
            wd01, _
            P8.FCDbl(wRs.Fields("手数料").Value), _
            P8.FCDbl(wRs.Fields("保証料").Value)

    wRs.MoveNext
    Loop
    End If
    wRs.Close
    Set wRs = Nothing
        
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_損益利息一覧表
'------------------------------------------------
Private Sub MX040_損益利息一覧表(pCsvFileName As String)
'
    Dim j As Integer
    Dim wd01 As Double
    Dim ws01 As String, ws02 As String, wsNendo As String
    Dim wsTbl As String, wstr2 As String
    Dim wWhere As String, wOrder As String
    Dim FLG_Order As Boolean
    
    Dim wsF01 As String, wsF02 As String

    Dim wDate1 As Date
    Dim w開始年月日 As Date
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    'GInt1パラメータ
    ws01 = Right("00" & CStr(GInt1), 2)

    If GRpt.推移 = "年次" Then
        'w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "平成")
        '2019/01/15 日付入力区分 仕様変更
        If G基本情報.日付入力区分 = "0" Then
        '和暦
            w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "平成")
        Else
        '西暦
            w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "西暦")
        End If
        
        wDate1 = DateAdd("d", -1, DateAdd("yyyy", 1, w開始年月日))
    Else
        GVar1 = C年月日.平成To西暦("年月", GRpt.テキスト_02)
        wDate1 = MBA010_締日年月日(CDate(GVar1))
    End If
'
    wstr = "Select "
    wstr = wstr & "K.借入番号 As 借入番号,"
    
    'セクションGR
    wstr = wstr & "K.銀行番号 As GrpFld_Ginko,"
    wstr = wstr & "G.金融機関番号 As GrpFld_Kinyu,"
    wstr = wstr & "B.部門番号 As GrpFld_Bumon,"
    wstr = wstr & " K.借入金種別区分 As GrpFld_KShubetu,"
    If GStr = "金利GR" Then
        wstr = wstr & "IIF(K.金利グループ区分<>'',K.金利グループ区分,'99999') As GrpFld_KGroup,"
    End If
    
    wstr = wstr & "G.銀行名 As 銀行名,"
    wstr = wstr & "G.金融機関名 As 金融機関名,"
    wstr = wstr & "B.部門名 As 部門名,"
    wstr = wstr & "S.借入金種別名 As 借入金種別名,"
    If GStr <> "金利GR" Then
        wstr = wstr & "'' As 金利グループ名,"
    Else
        wstr = wstr & "IIF(KG.金利グループ名<>'',KG.金利グループ名,'グループ無') As 金利グループ名,"
    End If
    
    wstr = wstr & " IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期','長期') As 長短区分," 'As 長短区分,"
    'wstr = wstr & "利率_" + ws01 + " As 利率,"
    wstr = wstr & " IIF(利率_" + ws01 + "<>0,利率_" + ws01 + ",利率合計) As 利率,"
    wstr = wstr & "利息_" + ws01 + " As 利息額,"

    wstr = wstr & "Z.前払利息_" & ws01 & " - Z.前払利息増_" & ws01 & " + Z.前払利息減_" & ws01 & " As 前前払利息残,"
    wstr = wstr & "Z.前払利息増_" & ws01 & " As 前払利息増,"
    wstr = wstr & "Z.前払利息減_" & ws01 & " As 前払利息減,"
    wstr = wstr & "Z.前払利息_" & ws01 & " As 当前払利息残,"

    wstr = wstr & "Z.未払利息_" & ws01 & " - Z.未払利息増_" & ws01 & " + Z.未払利息減_" & ws01 & " As 前未払利息残,"
    wstr = wstr & "Z.未払利息増_" & ws01 & " As 未払利息増,"
    wstr = wstr & "Z.未払利息減_" & ws01 & " As 未払利息減,"
    wstr = wstr & "Z.未払利息_" & ws01 & " As 当未払利息残,"

    wstr = wstr & "Z2.損益利息額_" & ws01 & " As 損益利息額,"

    wstr = wstr & "Z2.損益利息額_01"
    For j = 2 To GInt1
        ws02 = Right("00" + CStr(j), 2)
        wstr = wstr & " + Z2.損益利息額_" & ws02
    Next j
    wstr = wstr & " As 累計損益利息額"
        
    wstr = wstr & " FROM (((((DCDA010_借入残高推移表結果 As Z"
    wstr = wstr & " INNER JOIN DCDA010_借入残高推移表結果２ As Z2"
    wstr = wstr & " ON Z.借入番号=Z2.借入番号)"
    wstr = wstr & " INNER JOIN DCIA010_借入金ワーク As K"
    wstr = wstr & " ON Z.借入番号=K.借入番号)"
    wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr & " ON K.銀行番号=G.銀行番号)"
    wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
    wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分)"
    wstr = wstr & " LEFT JOIN DAAA200_部門マスタ As B"
    wstr = wstr & " ON K.プロジェクト番号 = B.部門番号)"
    
    If GStr = "金利GR" Then
        wstr = wstr & " LEFT JOIN DAAA115_金利シミュレーショングループ As KG"
        wstr = wstr & " ON K.金利グループ区分 = KG.金利グループ区分"
    End If

    wWhere = ""
    wWhere = " Where Format(K.実行日,'yyyymmdd') <= '" & Format(wDate1, "yyyymmdd") & "'"
    
'    wWhere = wWhere & " Where 融資_" + ws01 + "<>0"
'    wWhere = wWhere & " Or 元金_" + ws01 + "<>0"
'    wWhere = wWhere & " Or 利息_" + ws01 + "<>0"
'    wWhere = wWhere & " Or 返済_" + ws01 + "<>0"
'    wWhere = wWhere & " Or 解約_" + ws01 + "<>0"
'    wWhere = wWhere & " Or 残高_" + ws01 + "<>0"
'    wWhere = wWhere & " Or Z.前払利息_" & ws01 & "<>0"
'    wWhere = wWhere & " Or Z.前払利息増_" & ws01 & "<>0"
'    wWhere = wWhere & " Or Z.前払利息減_" & ws01 & "<>0"
'    wWhere = wWhere & " Or Z.前払利息_" & ws01 & "<>0"
'
'    wWhere = wWhere & " Or Z.未払利息_" & ws01 & "<>0"
'    wWhere = wWhere & " Or Z.未払利息増_" & ws01 & "<>0"
'    wWhere = wWhere & " Or Z.未払利息減_" & ws01 & "<>0"
'    wWhere = wWhere & " Or Z.未払利息_" & ws01 & "<>0"
'
'    wWhere = wWhere & " Or Z2.損益利息額_" & ws01 & "<>0"
    
    wstr = wstr & wWhere

     'Order
'    If GStr <> "金利GR" Then
'        wWhere = wWhere & " ORDER BY K.借入金種別区分,K.銀行番号,Z.借入番号"
'        'wWhere = wWhere & " ORDER By K.銀行番号,借入金種別区分,Z.借入番号"
'    Else
'        '金利SM
'        wWhere = wWhere & " ORDER BY IIF(K.金利グループ区分<>'',K.金利グループ区分,'99999'),K.銀行番号,Z.借入番号"
'    End If
    
    'wOrder
    wOrder = "": FLG_Order = False
    For j = 1 To 4
        If (GRpt.S_金融 = "分類" & CStr(j) Or GRpt.S_銀行 = "分類" & CStr(j)) _
        And FLG_Order = False Then
            wOrder = wOrder & "K.銀行番号,"
            FLG_Order = True
        ElseIf GRpt.S_種別 = "分類" & CStr(j) Then
            wOrder = wOrder & "K.借入金種別区分,"
        ElseIf GRpt.S_部門 = "分類" & CStr(j) Then
            wOrder = wOrder & "B.部門番号,"
        ElseIf GRpt.S_金利 = "分類" & CStr(j) Then
            If GStr <> "金利GR" Then
                wOrder = wOrder & "K.借入金種別区分,"
            Else
            '金利SM
                wOrder = wOrder & "IIF(K.金利グループ区分<>'',K.金利グループ区分,'99999'),"
            End If
        End If
    Next j
    wOrder = " Order by " & wOrder & "K.借入番号"
    wstr = wstr & wOrder
    
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        
        '名称
        Write #1, _
            "借入番号", _
            "長短区分", _
            "利率", _
            "利息額", _
            "前前払利息残", _
            "前払利息増", _
            "前払利息減", _
            "当前払利息残", _
            "前未払利息残", _
            "未払利息増", _
            "未払利息減", _
            "当未払利息残", _
            "損益利息額", _
            "累計損益利息額"
    
    Do Until wRs.eof
    
        Write #1, _
            P8.FCStr(wRs.Fields("借入番号").Value), _
            P8.FCStr(wRs.Fields("長短区分").Value), _
            P8.FCDbl(wRs.Fields("利率").Value), _
            P8.FCDbl(wRs.Fields("利息額").Value), _
            P8.FCDbl(wRs.Fields("前前払利息残").Value), _
            P8.FCDbl(wRs.Fields("前払利息増").Value), _
            P8.FCDbl(wRs.Fields("前払利息減").Value), _
            P8.FCDbl(wRs.Fields("当前払利息残").Value), _
            P8.FCDbl(wRs.Fields("前未払利息残").Value), _
            P8.FCDbl(wRs.Fields("未払利息増").Value), _
            P8.FCDbl(wRs.Fields("未払利息減").Value), _
            P8.FCDbl(wRs.Fields("当未払利息残").Value), _
            P8.FCDbl(wRs.Fields("損益利息額").Value), _
            P8.FCDbl(wRs.Fields("累計損益利息額").Value)

    wRs.MoveNext
    Loop
    End If
    wRs.Close
    Set wRs = Nothing
    
    '合計
    wstr2 = "select "
    wstr2 = wstr2 & "'総合計' As タイトル,'',count(Z.借入番号) & '件' As 件数,'','',"
    wstr2 = wstr2 & "'','','','','','','','','','','','','','','','','','','','','','','','','','','',"

    wstr2 = wstr2 & "Sum(利息_" & ws01 & ") As 利息額,"
    wstr2 = wstr2 & "Sum(Z.前払利息_" & ws01 & " - Z.前払利息増_" & ws01 & " + Z.前払利息減_" & ws01 & ") As 前前払利息残,"
    wstr2 = wstr2 & "Sum(前払利息増_" & ws01 & ") As 前払利息増,"
    wstr2 = wstr2 & "Sum(前払利息減_" & ws01 & ") As 前払利息減,"
    wstr2 = wstr2 & "Sum(前払利息_" & ws01 & ") As 当前払利息残,"
    wstr2 = wstr2 & "Sum(Z.未払利息_" & ws01 & " - Z.未払利息増_" & ws01 & " + Z.未払利息減_" & ws01 & ") As 前未払利息残,"
    wstr2 = wstr2 & "Sum(未払利息増_" & ws01 & ") As 未払利息増,"
    wstr2 = wstr2 & "Sum(未払利息減_" & ws01 & ") As 未払利息減,"
    wstr2 = wstr2 & "Sum(未払利息_" & ws01 & ") As 当未払利息残,"
    wstr2 = wstr2 & "Sum(損益利息額_" & ws01 & ") As 損益利息額,"
    
    wstr2 = wstr2 & "Sum(Z2.損益利息額_01"
    For j = 2 To GInt1
        ws02 = Right("00" + CStr(j), 2)
        wstr2 = wstr2 & " + Z2.損益利息額_" & ws02
    Next j
    wstr2 = wstr2 & ") As 累計損益利息額"
    
    wstr2 = wstr2 & " FROM (((((DCDA010_借入残高推移表結果 As Z"
    wstr2 = wstr2 & " INNER JOIN DCDA010_借入残高推移表結果２ As Z2"
    wstr2 = wstr2 & " ON Z.借入番号=Z2.借入番号)"
    wstr2 = wstr2 & " INNER JOIN DCIA010_借入金ワーク As K"
    wstr2 = wstr2 & " ON Z.借入番号=K.借入番号)"
    wstr2 = wstr2 & " INNER JOIN DAAA040_銀行マスタ As G"
    wstr2 = wstr2 & " ON K.銀行番号=G.銀行番号)"
    wstr2 = wstr2 & " LEFT JOIN DAAA116_借入金種別 As S"
    wstr2 = wstr2 & " ON K.借入金種別区分 = S.借入金種別区分)"
    wstr2 = wstr2 & " LEFT JOIN DAAA200_部門マスタ As B"
    wstr2 = wstr2 & " ON K.プロジェクト番号 = B.部門番号)"
    
    If GStr = "金利GR" Then
        wstr = wstr & " LEFT JOIN DAAA115_金利シミュレーショングループ As KG"
        wstr = wstr & " ON K.金利グループ区分 = KG.金利グループ区分"
    End If

    wWhere = ""
    wWhere = " Where Format(K.実行日,'yyyymmdd') <= '" & Format(wDate1, "yyyymmdd") & "'"
    
'    wWhere = wWhere & " Where 融資_" + ws01 + "<>0"
'    wWhere = wWhere & " Or 元金_" + ws01 + "<>0"
'    wWhere = wWhere & " Or 利息_" + ws01 + "<>0"
'    wWhere = wWhere & " Or 返済_" + ws01 + "<>0"
'    wWhere = wWhere & " Or 解約_" + ws01 + "<>0"
'    wWhere = wWhere & " Or 残高_" + ws01 + "<>0"
'    wWhere = wWhere & " Or Z.前払利息_" & ws01 & "<>0"
'    wWhere = wWhere & " Or Z.前払利息増_" & ws01 & "<>0"
'    wWhere = wWhere & " Or Z.前払利息減_" & ws01 & "<>0"
'    wWhere = wWhere & " Or Z.前払利息_" & ws01 & "<>0"
'
'    wWhere = wWhere & " Or Z.未払利息_" & ws01 & "<>0"
'    wWhere = wWhere & " Or Z.未払利息増_" & ws01 & "<>0"
'    wWhere = wWhere & " Or Z.未払利息減_" & ws01 & "<>0"
'    wWhere = wWhere & " Or Z.未払利息_" & ws01 & "<>0"
'
'    wWhere = wWhere & " Or Z2.損益利息額_" & ws01 & "<>0"
    
    wstr2 = wstr2 & wWhere
    
    Call AdoRecordsetOpen(GDb, wRs, wstr2)
    If Not wRs.eof Then
    Do Until wRs.eof
    
        Write #1, _
            P8.FCStr(wRs.Fields("タイトル").Value), "", P8.FCStr(wRs.Fields("件数").Value), _
            P8.FCDbl(wRs.Fields("利息額").Value), _
            P8.FCDbl(wRs.Fields("前前払利息残").Value), _
            P8.FCDbl(wRs.Fields("前払利息増").Value), _
            P8.FCDbl(wRs.Fields("前払利息減").Value), _
            P8.FCDbl(wRs.Fields("当前払利息残").Value), _
            P8.FCDbl(wRs.Fields("前未払利息残").Value), _
            P8.FCDbl(wRs.Fields("未払利息増").Value), _
            P8.FCDbl(wRs.Fields("未払利息減").Value), _
            P8.FCDbl(wRs.Fields("当未払利息残").Value), _
            P8.FCDbl(wRs.Fields("損益利息額").Value), _
            P8.FCDbl(wRs.Fields("累計損益利息額").Value)

    wRs.MoveNext
    Loop
    End If
    wRs.Close
    Set wRs = Nothing
        
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_前払利息残高表
'------------------------------------------------
Private Sub MX040_前払利息残高表(pCsvFileName As String)
'
    Dim j As Integer
    Dim wd01 As Double
    Dim w番号 As String
    Dim ws01 As String
    Dim wsF01 As String, wsF02 As String
'
    On Error Resume Next
'
    If GRpt.推移 = "月次" Then
         wsF01 = "前月前払利息残高"
         wsF02 = "当月前払利息残高"
    Else
         wsF01 = "前期前払利息残高"
         wsF02 = "当期前払利息残高"
    End If
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    'GInt1パラメータ
    w番号 = Right("00" & CStr(GInt1), 2)
    
    wstr = ""
    wstr = wstr & "Select "
    wstr = wstr & "K.銀行番号,"
    wstr = wstr & "G.銀行名,"
    wstr = wstr + "B.部門名,"
    wstr = wstr & "K.借入番号,"
    wstr = wstr & "K.借入内容,"
    'wstr = wstr & "format(K.利率,'#,##0.00000') As 利率,"
    wstr = wstr & "format(Z.利率_" & w番号 & ",'#,##0.00000') As 利率,"
    wstr = wstr & "IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As 金利種別,"
    wstr = wstr & "KK.基準金利名,"
    wstr = wstr & "K.金利条件,"
    wstr = wstr & "IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期借入金','長期借入金') AS 長短区分,"
    wstr = wstr & "IIF(K.有担保フラグ = " & P8.FCDbl(XMXA020_区分("有担フラグ", "無担保")) & ",'無担保','有担保') As 担保区分,"
    
    wstr = wstr & "Z.残高_" & w番号 & " As 融資残高,"
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & "-Z.前払利息増_" & w番号 & "+Z.前払利息_" & w番号 & ",-Z.未払利息増_" & w番号 & "+Z.未払利息減_" & w番号 & "+Z.未払利息_" & w番号 & ") As 前利息残高,"
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息増_" & w番号 & ",Z.未払利息増_" & w番号 & ") As 利息増,"
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & ",Z.未払利息減_" & w番号 & ") As 利息減,"
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息_" & w番号 & ",Z.未払利息_" & w番号 & ") As 利息残高,"
    wstr = wstr & "Z2.損益利息額_" & w番号 & " As 損益利息額,"
    
    wstr = wstr & "SH.支払区分名 As 支払日,"
    wstr = wstr & "IIF(K.営業日区分 = " & P8.FCDbl(XMXA020_区分("営業日", "翌営業日")) & ",'翌営業日','前営業日') As 営業日,"
    wstr = wstr & "IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As 利息区分,"
    wstr = wstr & "IIF(K.利息計算日数区分 = " & P8.FCDbl(XMXA020_区分("利息計算日数", "営業日数")) & ",'営業日数','固定日数') As 利息日数,"
    wstr = wstr & "IIF(K.利息支払方法 = " & P8.FCDbl(XMXA020_区分("利息支払", "毎月")) & ",'毎月','一括') As 利息支払方法,"
    wstr = wstr & "IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "控除無し")) & ",'控除無し',"
    wstr = wstr & "IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "実行日控除")) & ",'実行日控除',"
    wstr = wstr & "IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "最終返済日控除")) & ",'最終返済日控除',"
    wstr = wstr & "IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) & ",'実行日及び最終返済日控除','中間利払最終日控除')))) As 利息控除区分,"
    wstr = wstr & "IIF(K.金利計算年間日数 = " & P8.FCDbl(XMXA020_区分("金利計算", "365日")) & ",'365','360') As 金利年間日数,"
    
    wstr = wstr & "K.借入金種別区分,"
    wstr = wstr & "S.借入金種別名,"
    wstr = wstr & "K.金利グループ区分,"
    wstr = wstr & "IIF(KG.金利グループ名<>'',KG.金利グループ名,'グループ無') As 金利グループ名"
    
    wstr = wstr & " FROM (((((((DCDA010_借入残高推移表結果 As Z"
    wstr = wstr & " INNER JOIN DCDA010_借入残高推移表結果２ As Z2"
    wstr = wstr & " ON Z.借入番号=Z2.借入番号)"
    wstr = wstr & " INNER JOIN DBDA010_借入金 As K"
    wstr = wstr & " ON Z.借入番号=K.借入番号)"
    wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr & " ON K.銀行番号=G.銀行番号)"
    wstr = wstr + " Inner Join DAAB020_支払区分マスタ As SH"
    wstr = wstr + " ON K.支払日 = SH.支払日)"
    wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
    wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分)"
    wstr = wstr + " LEFT JOIN DAAA200_部門マスタ AS B"
    wstr = wstr + "  ON B.部門番号 = K.プロジェクト番号)"
    wstr = wstr + " Left Join DAAA116_基準金利 As KK"
    wstr = wstr + " ON K.基準金利区分 = KK.基準金利区分)"
    wstr = wstr & " LEFT JOIN DAAA115_金利シミュレーショングループ As KG"
    wstr = wstr & " ON K.金利グループ区分 = KG.金利グループ区分"
    
    'Where 条件
    'All 0 は表示しない
    wstr = wstr & " Where K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "'"
    wstr = wstr & " And (Z.前払利息増_" & w番号 & "<>0"
    wstr = wstr & " Or Z.前払利息減_" & w番号 & "<>0"
    wstr = wstr & " Or Z.前払利息_" & w番号 & "<>0"
    wstr = wstr & " Or Z2.損益利息額_" & w番号 & "<>0)"
    'wstr = wstr & " Or Z.未払利息増_" & w番号 & "<>0"
    'wstr = wstr & " Or Z.未払利息減_" & w番号 & "<>0"
    'wstr = wstr & " Or Z.未払利息_" & w番号 & "<>0"
    wstr = wstr & " ORDER BY K.借入金種別区分,K.銀行番号,K.借入番号"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        '名称
        Write #1, _
            "銀行番号", "銀行名", "部門名", "借入番号", "借入内容", _
            "利率", "金利種別", "基準金利名", "金利条件", "長短区分", "担保区分", _
            "融資残高", wsF01, "前払利息増", "前払利息減", wsF02, "損益利息額", _
            "支払日", "営業日", "利息区分", "利息日数", "利息支払方法", "利息控除区分", "金利年間日数", _
            "借入金種別区分", "借入金種別名", "金利グループ区分", "金利グループ名"
    
    Do Until wRs.eof
    
        Write #1, _
            P8.FCStr(wRs.Fields("銀行番号").Value), P8.FCStr(wRs.Fields("銀行名").Value), P8.FCStr(wRs.Fields("部門名").Value), _
            P8.FCStr(wRs.Fields("借入番号").Value), P8.FCStr(wRs.Fields("借入内容").Value), _
            P8.FCDbl(wRs.Fields("利率").Value), _
            P8.FCStr(wRs.Fields("金利種別").Value), P8.FCStr(wRs.Fields("基準金利名").Value), P8.FCStr(wRs.Fields("金利条件").Value), _
            P8.FCStr(wRs.Fields("長短区分").Value), P8.FCStr(wRs.Fields("担保区分").Value), _
            P8.FCDbl(wRs.Fields("融資残高").Value), P8.FCDbl(wRs.Fields("前利息残高").Value), _
            P8.FCDbl(wRs.Fields("利息増").Value), P8.FCDbl(wRs.Fields("利息減").Value), _
            P8.FCDbl(wRs.Fields("利息残高").Value), _
            P8.FCDbl(wRs.Fields("損益利息額").Value), _
            P8.FCStr(wRs.Fields("支払日").Value), P8.FCStr(wRs.Fields("営業日").Value), _
            P8.FCStr(wRs.Fields("利息区分").Value), P8.FCStr(wRs.Fields("利息日数").Value), _
            P8.FCStr(wRs.Fields("利息支払方法").Value), P8.FCStr(wRs.Fields("利息控除区分").Value), _
            P8.FCStr(wRs.Fields("金利年間日数").Value), _
            P8.FCStr(wRs.Fields("借入金種別区分").Value), P8.FCStr(wRs.Fields("借入金種別名").Value), _
            P8.FCStr(wRs.Fields("金利グループ区分").Value), P8.FCStr(wRs.Fields("金利グループ名").Value)
    
    wRs.MoveNext
    Loop
    End If
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_未払利息残高表
'------------------------------------------------
Private Sub MX040_未払利息残高表(pCsvFileName As String)
'
    Dim j As Integer
    Dim wd01 As Double
    Dim w番号 As String
    Dim ws01 As String
    Dim wsF01 As String, wsF02 As String
'
    On Error Resume Next
'
    If GRpt.推移 = "月次" Then
         wsF01 = "前月未払利息残高"
         wsF02 = "当月未払利息残高"
    Else
         wsF01 = "前期未払利息残高"
         wsF02 = "当期未払利息残高"
    End If
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    'GInt1パラメータ
    w番号 = Right("00" & CStr(GInt1), 2)
    
    wstr = ""
    wstr = wstr & "Select "
    wstr = wstr & "K.銀行番号,"
    wstr = wstr & "G.銀行名,"
    wstr = wstr + "B.部門名,"
    wstr = wstr & "K.借入番号,"
    wstr = wstr & "K.借入内容,"
    'wstr = wstr & "format(K.利率,'#,##0.00000') As 利率,"
    wstr = wstr & "format(Z.利率_" & w番号 & ",'#,##0.00000') As 利率,"
    wstr = wstr + "IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As 金利種別,"
    wstr = wstr + "KK.基準金利名,"
    wstr = wstr & "K.金利条件,"
    wstr = wstr & "IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期借入金','長期借入金') AS 長短区分,"
    wstr = wstr & "IIF(K.有担保フラグ = " & P8.FCDbl(XMXA020_区分("有担フラグ", "無担保")) & ",'無担保','有担保') As 担保区分,"
    
    wstr = wstr & "Z.残高_" & w番号 & " As 融資残高,"
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & "-Z.前払利息増_" & w番号 & "+Z.前払利息_" & w番号 & ",-Z.未払利息増_" & w番号 & "+Z.未払利息減_" & w番号 & "+Z.未払利息_" & w番号 & ") As 前利息残高,"
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息増_" & w番号 & ",Z.未払利息増_" & w番号 & ") As 利息増,"
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & ",Z.未払利息減_" & w番号 & ") As 利息減,"
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息_" & w番号 & ",Z.未払利息_" & w番号 & ") As 利息残高,"
    wstr = wstr & "Z2.損益利息額_" & w番号 & " As 損益利息額,"
    
    wstr = wstr & "SH.支払区分名 As 支払日,"
    wstr = wstr & "IIF(K.営業日区分 = " & P8.FCDbl(XMXA020_区分("営業日", "翌営業日")) & ",'翌営業日','前営業日') As 営業日,"
    wstr = wstr & "IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As 利息区分,"
    wstr = wstr & "IIF(K.利息計算日数区分 = " & P8.FCDbl(XMXA020_区分("利息計算日数", "営業日数")) & ",'営業日数','固定日数') As 利息日数,"
    wstr = wstr & "IIF(K.利息支払方法 = " & P8.FCDbl(XMXA020_区分("利息支払", "毎月")) & ",'毎月','一括') As 利息支払方法,"
    wstr = wstr & "IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "控除無し")) & ",'控除無し',"
    wstr = wstr & "IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "実行日控除")) & ",'実行日控除',"
    wstr = wstr & "IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "最終返済日控除")) & ",'最終返済日控除',"
    wstr = wstr & "IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) & ",'実行日及び最終返済日控除','中間利払最終日控除')))) As 利息控除区分,"
    wstr = wstr & "IIF(K.金利計算年間日数 = " & P8.FCDbl(XMXA020_区分("金利計算", "365日")) & ",'365','360') As 金利年間日数,"
    
    wstr = wstr & "K.借入金種別区分,"
    wstr = wstr & "S.借入金種別名,"
    wstr = wstr & "K.金利グループ区分,"
    wstr = wstr & "IIF(KG.金利グループ名<>'',KG.金利グループ名,'グループ無') As 金利グループ名"
    
    wstr = wstr & " FROM (((((((DCDA010_借入残高推移表結果 As Z"
    wstr = wstr & " INNER JOIN DCDA010_借入残高推移表結果２ As Z2"
    wstr = wstr & " ON Z.借入番号=Z2.借入番号)"
    wstr = wstr & " INNER JOIN DBDA010_借入金 As K"
    wstr = wstr & " ON Z.借入番号=K.借入番号)"
    wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr & " ON K.銀行番号=G.銀行番号)"
    wstr = wstr + " Inner Join DAAB020_支払区分マスタ As SH"
    wstr = wstr + " ON K.支払日 = SH.支払日)"
    wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
    wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分)"
    wstr = wstr + " LEFT JOIN DAAA200_部門マスタ AS B"
    wstr = wstr + "  ON B.部門番号 = K.プロジェクト番号)"
    wstr = wstr + " Left Join DAAA116_基準金利 As KK"
    wstr = wstr + " ON K.基準金利区分 = KK.基準金利区分)"
    wstr = wstr & " LEFT JOIN DAAA115_金利シミュレーショングループ As KG"
    wstr = wstr & " ON K.金利グループ区分 = KG.金利グループ区分"
    
    'Where 条件
    'All 0 は表示しない
    wstr = wstr & " Where K.利息区分='" & XMXA020_区分("利息区分", "利息後払") & "'"
    wstr = wstr & " And (Z.未払利息増_" & w番号 & "<>0"
    wstr = wstr & " Or Z.未払利息減_" & w番号 & "<>0"
    wstr = wstr & " Or Z.未払利息_" & w番号 & "<>0"
    wstr = wstr & " Or Z2.損益利息額_" & w番号 & "<>0)"
    wstr = wstr & " ORDER BY K.借入金種別区分,K.利息区分,K.銀行番号,K.借入番号"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        '名称
        Write #1, _
            "銀行番号", "銀行名", "部門名", "借入番号", "借入内容", _
            "利率", "金利種別", "基準金利名", "金利条件", "長短区分", "担保区分", _
            "融資残高", wsF01, "未払利息増", "未払利息減", wsF02, "損益利息額", _
            "支払日", "営業日", "利息区分", "利息日数", "利息支払方法", "利息控除区分", "金利年間日数", _
            "借入金種別区分", "借入金種別名", "金利グループ区分", "金利グループ名"
    
    Do Until wRs.eof
    
        Write #1, _
            P8.FCStr(wRs.Fields("銀行番号").Value), P8.FCStr(wRs.Fields("銀行名").Value), P8.FCStr(wRs.Fields("部門名").Value), _
            P8.FCStr(wRs.Fields("借入番号").Value), P8.FCStr(wRs.Fields("借入内容").Value), _
            P8.FCDbl(wRs.Fields("利率").Value), _
            P8.FCStr(wRs.Fields("金利種別").Value), P8.FCStr(wRs.Fields("基準金利名").Value), P8.FCStr(wRs.Fields("金利条件").Value), _
            P8.FCStr(wRs.Fields("長短区分").Value), P8.FCStr(wRs.Fields("担保区分").Value), _
            P8.FCDbl(wRs.Fields("融資残高").Value), P8.FCDbl(wRs.Fields("前利息残高").Value), _
            P8.FCDbl(wRs.Fields("利息増").Value), P8.FCDbl(wRs.Fields("利息減").Value), _
            P8.FCDbl(wRs.Fields("利息残高").Value), _
            P8.FCDbl(wRs.Fields("損益利息額").Value), _
            P8.FCStr(wRs.Fields("支払日").Value), P8.FCStr(wRs.Fields("営業日").Value), _
            P8.FCStr(wRs.Fields("利息区分").Value), P8.FCStr(wRs.Fields("利息日数").Value), _
            P8.FCStr(wRs.Fields("利息支払方法").Value), P8.FCStr(wRs.Fields("利息控除区分").Value), _
            P8.FCStr(wRs.Fields("金利年間日数").Value), _
            P8.FCStr(wRs.Fields("借入金種別区分").Value), P8.FCStr(wRs.Fields("借入金種別名").Value), _
            P8.FCStr(wRs.Fields("金利グループ区分").Value), P8.FCStr(wRs.Fields("金利グループ名").Value)

    wRs.MoveNext
    Loop
    End If
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_前払利息残高推移表
'------------------------------------------------
Private Sub MX040_前払利息残高推移表(pCsvFileName As String)
'
    Dim j As Integer, wML As Integer
    Dim ws01 As String, wsNendo As String
    Dim wsTbl As String
'
    On Error Resume Next
'
    'フィールド数
    Select Case GRpt.推移
    Case "月次", "四半期"
        wML = 12
    Case "年次", "半期"
        wML = 10
    End Select

    wsTbl = "DBDA010_借入金"
    If GRpt.帳票名 = "借入利息残高推移表" Then
        wsTbl = "DBDA010_借入金"
    ElseIf GRpt.帳票名 = "貸付利息残高推移表" Then
        wsTbl = "DBDA010_貸付金"
    End If
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wstr = "SELECT "
    
    If GRpt.帳票名 = "借入利息残高推移表" Then
        wstr = wstr & "Z.借入番号,"
    ElseIf GRpt.帳票名 = "貸付利息残高推移表" Then
        wstr = wstr & "Z.借入番号 As 貸付番号,"
    End If

    wstr = wstr & " K.借入内容,"
    wstr = wstr & " KS.借入金種別名,"
    wstr = wstr & " K.銀行番号,"
    wstr = wstr & " G.銀行名,"
    wstr = wstr + "B.部門名,"
    
    wstr = wstr & "IIF(K.営業日区分 = " & P8.FCDbl(XMXA020_区分("営業日", "翌営業日")) & ",'翌営業日','前営業日') As 営業日,"
    wstr = wstr & "IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As 利息区分,"
    wstr = wstr & "IIF(K.利息計算日数区分 = " & P8.FCDbl(XMXA020_区分("利息計算日数", "営業日数")) & ",'営業日数','固定日数') As 利息計算日数,"
    wstr = wstr & "IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As 金利種別,"
    wstr = wstr & "KK.基準金利名,"
    wstr = wstr & "K.金利条件 As 金利備考,"
    
    wstr = wstr & "IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期借入金','長期借入金') AS 長短区分,"
    wstr = wstr & "IIF(K.有担保フラグ = " & P8.FCDbl(XMXA020_区分("有担フラグ", "無担保")) & ",'無担保','有担保') As 担保区分,"
    wstr = wstr & "K.担保名 As 担保内容,"
    wstr = wstr & "IIF(設備フラグ=1,'設備','運転資金') As 設備区分,"
    wstr = wstr & "K.資金用途,"
    
    wstr = wstr & "KG.金利グループ名,"
    wstr = wstr & "IIF(K.sm区分=1 and K.金融リストラ番号<>'','借入SM',IIF(K.sm区分=0 and not (K.金融解約実行日 is null),'解約SM','')) As ｼﾐｭﾚｰｼｮﾝ内容,"
    wstr = wstr & "K.金融リストラ番号 As ｼﾐｭﾚｰｼｮﾝ番号,"
    'wstr = wstr & "format(K.金融解約実行日,'yyyy/mm/dd') As ｼﾐｭﾚｰｼｮﾝ解約日,"
    
    'wstr = wstr & "Format(K.実行日,'yyyy/mm/dd') As 実行日,"
    'wstr = wstr & "Format(K.初回返済年月,'yyyy/mm/dd') As 初回返済年月,"
    'wstr = wstr & "Format(K.初回返済実行日,'yyyy/mm/dd') As 初回返済年月日,"
    'wstr = wstr & "Format(K.最終返済年月,'yyyy/mm/dd') As 最終返済年月,"
    'wstr = wstr & "Format(K.最終返済実行日,'yyyy/mm/dd') As 最終返済年月日,"
    'wstr = wstr & "Format(K.解約実行日,'yyyy/mm/dd') As 解約年月日,"
    
    '2012.10.23　追加 by k.kunita
    wstr = wstr & "format(K.金融解約実行日,'" & Gfmtcsv年月日 & "') As ｼﾐｭﾚｰｼｮﾝ解約日,"
    
    wstr = wstr & "Format(K.実行日,'" & Gfmtcsv年月日 & "') As 実行日,"
    wstr = wstr & "Format(K.初回返済年月,'" & Gfmtcsv年月日 & "') As 初回返済年月,"
    wstr = wstr & "Format(K.初回返済実行日,'" & Gfmtcsv年月日 & "') As 初回返済年月日,"
    wstr = wstr & "Format(K.最終返済年月,'" & Gfmtcsv年月日 & "') As 最終返済年月,"
    wstr = wstr & "Format(K.最終返済実行日,'" & Gfmtcsv年月日 & "') As 最終返済年月日,"
    wstr = wstr & "Format(K.解約実行日,'" & Gfmtcsv年月日 & "') As 解約年月日,"
    
    wstr = wstr & "Format(IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.初回返済額,0),'#,##0') As 初回返済額,"
    wstr = wstr & "Format(IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.毎月返済額,0),'#,##0') As 毎月返済額,"
    wstr = wstr & "Format(IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.最終返済額,0),'#,##0') As 最終返済額,"
    
    wstr = wstr & "K.返済単位月数,"
    
    wstr = wstr & "Format(残高合計,'#,##0') As 合計融資残高,"
    wstr = wstr & "Format(前払利息増合計,'#,##0') As 合計前払利息増,"
    wstr = wstr & "Format(前払利息減合計,'#,##0') As 合計前払利息減,"
    wstr = wstr & "Format(前払利息合計,'#,##0') As 合計前払利息残高,"
    wstr = wstr & "Format(損益利息額合計,'#,##0') As 合計損益利息額,"
    
    For j = 1 To wML - 1
        ws01 = "_" & Right("00" & CStr(j), 2)
        
        If G基本情報.日付入力区分 = "0" Then
        '和暦入力
            wstr = wstr & "Format(残高" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "融資残高,"
            wstr = wstr & "Format(前払利息増" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "合計前払利息増,"
            wstr = wstr & "Format(前払利息減" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "合計前払利息減,"
            wstr = wstr & "Format(前払利息" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "合計前払利息残高,"
            wstr = wstr & "Format(損益利息額" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "合計損益利息額,"
        Else
        '西暦入力
            wstr = wstr & "Format(残高" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "融資残高,"
            wstr = wstr & "Format(前払利息増" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "合計前払利息増,"
            wstr = wstr & "Format(前払利息減" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "合計前払利息減,"
            wstr = wstr & "Format(前払利息" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "合計前払利息残高,"
            wstr = wstr & "Format(損益利息額" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "合計損益利息額,"
        End If
    Next j
    
    For j = wML To wML
        ws01 = "_" & Right("00" & CStr(j), 2)
        
        If G基本情報.日付入力区分 = "0" Then
        '和暦入力
            wstr = wstr & "Format(残高" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "融資残高,"
            wstr = wstr & "Format(前払利息増" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "合計前払利息増,"
            wstr = wstr & "Format(前払利息減" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "合計前払利息減,"
            wstr = wstr & "Format(前払利息" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "合計前払利息残高,"
            wstr = wstr & "Format(損益利息額" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "合計損益利息額"
        Else
        '西暦入力
            wstr = wstr & "Format(残高" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "融資残高,"
            wstr = wstr & "Format(前払利息増" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "合計前払利息増,"
            wstr = wstr & "Format(前払利息減" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "合計前払利息減,"
            wstr = wstr & "Format(前払利息" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "合計前払利息残高,"
            wstr = wstr & "Format(損益利息額" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "合計損益利息額"
        End If
    Next j

    wstr = wstr & " INTO [Text;database=" & wCsvDir & "]." & "[" & pCsvFileName & "]"
    
    wstr = wstr + " FROM ((((((DCDA010_借入残高推移表結果 As Z"
    wstr = wstr + " Inner Join DCDA010_借入残高推移表結果２ As Z2"
    wstr = wstr + " ON Z.借入番号 = Z2.借入番号)"
    wstr = wstr + " Inner Join " & wsTbl & " As K"
    wstr = wstr + " ON Z.借入番号 = K.借入番号)"
    wstr = wstr + " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr + " ON K.銀行番号 = G.銀行番号)"
    wstr = wstr + " Inner Join DAAA116_借入金種別 As KS"
    wstr = wstr + " ON K.借入金種別区分 = KS.借入金種別区分)"
    wstr = wstr + " LEFT JOIN DAAA200_部門マスタ AS B"
    wstr = wstr + "  ON B.部門番号 = K.プロジェクト番号)"
    wstr = wstr + " Left Join DAAA116_基準金利 As KK"
    wstr = wstr + " ON K.基準金利区分 = KK.基準金利区分)"
    wstr = wstr + " Left Join DAAA115_金利シミュレーショングループ As KG"
    wstr = wstr + " ON K.金利グループ区分 = KG.金利グループ区分"
    
    wstr = wstr & " Where K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "'"
    wstr = wstr & " And (前払利息増合計<>0"
    wstr = wstr & " Or 前払利息減合計<>0"
    wstr = wstr & " Or 前払利息合計<>0"
    wstr = wstr & " Or 損益利息額合計<>0"
     
    For j = 1 To wML
        ws01 = "_" & Right("00" & CStr(j), 2)
        
        wstr = wstr & " Or 前払利息増" & ws01 & "<>0"
        wstr = wstr & " Or 前払利息減" & ws01 & "<>0"
        wstr = wstr & " Or 前払利息" & ws01 & "<>0"
        wstr = wstr & " Or 損益利息額" & ws01 & "<>0"
    Next j

    wstr = wstr & " )"
    wstr = wstr & " Order by K.銀行番号,K.金利種別,K.有担保フラグ,Z.借入番号"
    
    GDb.Execute wstr
    
    DoEvents
    
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
End Sub

'------------------------------------------------
' MX040_未払利息残高推移表
'------------------------------------------------
Private Sub MX040_未払利息残高推移表(pCsvFileName As String)
'
    Dim j As Integer, wML As Integer
    Dim ws01 As String, wsNendo As String
    Dim wsTbl As String
'
    On Error Resume Next
'
    'フィールド数
    Select Case GRpt.推移
    Case "月次", "四半期"
        wML = 12
    Case "年次", "半期"
        wML = 10
    End Select

    wsTbl = "DBDA010_借入金"
    If GRpt.帳票名 = "借入利息残高推移表" Then
        wsTbl = "DBDA010_借入金"
    ElseIf GRpt.帳票名 = "貸付利息残高推移表" Then
        wsTbl = "DBDA010_貸付金"
    End If
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wstr = "SELECT "
    
    If GRpt.帳票名 = "借入利息残高推移表" Then
        wstr = wstr & "Z.借入番号,"
    ElseIf GRpt.帳票名 = "貸付利息残高推移表" Then
        wstr = wstr & "Z.借入番号 As 貸付番号,"
    End If

    wstr = wstr & " K.借入内容,"
    wstr = wstr & " KS.借入金種別名,"
    wstr = wstr & " K.銀行番号,"
    wstr = wstr & " G.銀行名,"
    wstr = wstr + "B.部門名,"
    
    wstr = wstr & "IIF(K.営業日区分 = " & P8.FCDbl(XMXA020_区分("営業日", "翌営業日")) & ",'翌営業日','前営業日') As 営業日,"
    wstr = wstr & "IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As 利息区分,"
    wstr = wstr & "IIF(K.利息計算日数区分 = " & P8.FCDbl(XMXA020_区分("利息計算日数", "営業日数")) & ",'営業日数','固定日数') As 利息計算日数,"
    wstr = wstr & "IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As 金利種別,"
    wstr = wstr & "KK.基準金利名,"
    wstr = wstr & "K.金利条件 As 金利備考,"
    
    wstr = wstr & "IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期借入金','長期借入金') AS 長短区分,"
    wstr = wstr & "IIF(K.有担保フラグ = " & P8.FCDbl(XMXA020_区分("有担フラグ", "無担保")) & ",'無担保','有担保') As 担保区分,"
    wstr = wstr & "K.担保名 As 担保内容,"
    wstr = wstr & "IIF(設備フラグ=1,'設備','運転資金') As 設備区分,"
    wstr = wstr & "K.資金用途,"
    
    wstr = wstr & "KG.金利グループ名,"
    wstr = wstr & "IIF(K.sm区分=1 and K.金融リストラ番号<>'','借入SM',IIF(K.sm区分=0 and not (K.金融解約実行日 is null),'解約SM','')) As ｼﾐｭﾚｰｼｮﾝ内容,"
    wstr = wstr & "K.金融リストラ番号 As ｼﾐｭﾚｰｼｮﾝ番号,"
    'wstr = wstr & "format(K.金融解約実行日,'yyyy/mm/dd') As ｼﾐｭﾚｰｼｮﾝ解約日,"
    
    'wstr = wstr & "Format(K.実行日,'yyyy/mm/dd') As 実行日,"
    'wstr = wstr & "Format(K.初回返済年月,'yyyy/mm/dd') As 初回返済年月,"
    'wstr = wstr & "Format(K.初回返済実行日,'yyyy/mm/dd') As 初回返済年月日,"
    'wstr = wstr & "Format(K.最終返済年月,'yyyy/mm/dd') As 最終返済年月,"
    'wstr = wstr & "Format(K.最終返済実行日,'yyyy/mm/dd') As 最終返済年月日,"
    'wstr = wstr & "Format(K.解約実行日,'yyyy/mm/dd') As 解約年月日,"
    
    '2012.10.23　追加 by k.kunita
    wstr = wstr & "format(K.金融解約実行日,'" & Gfmtcsv年月日 & "') As ｼﾐｭﾚｰｼｮﾝ解約日,"
    
    wstr = wstr & "Format(K.実行日,'" & Gfmtcsv年月日 & "') As 実行日,"
    wstr = wstr & "Format(K.初回返済年月,'" & Gfmtcsv年月日 & "') As 初回返済年月,"
    wstr = wstr & "Format(K.初回返済実行日,'" & Gfmtcsv年月日 & "') As 初回返済年月日,"
    wstr = wstr & "Format(K.最終返済年月,'" & Gfmtcsv年月日 & "') As 最終返済年月,"
    wstr = wstr & "Format(K.最終返済実行日,'" & Gfmtcsv年月日 & "') As 最終返済年月日,"
    wstr = wstr & "Format(K.解約実行日,'" & Gfmtcsv年月日 & "') As 解約年月日,"
    
    wstr = wstr & "Format(IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.初回返済額,0),'#,##0') As 初回返済額,"
    wstr = wstr & "Format(IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.毎月返済額,0),'#,##0') As 毎月返済額,"
    wstr = wstr & "Format(IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.最終返済額,0),'#,##0') As 最終返済額,"
    
    wstr = wstr & "K.返済単位月数,"

    wstr = wstr & "Format(残高合計,'#,##0') As 合計融資残高,"
    wstr = wstr & "Format(未払利息増合計,'#,##0') As 合計未払利息増,"
    wstr = wstr & "Format(未払利息減合計,'#,##0') As 合計未払利息減,"
    wstr = wstr & "Format(未払利息合計,'#,##0') As 合計未払利息残高,"
    wstr = wstr & "Format(損益利息額合計,'#,##0') As 合計損益利息額,"
    
    For j = 1 To wML - 1
        ws01 = "_" & Right("00" & CStr(j), 2)
        
        If G基本情報.日付入力区分 = "0" Then
        '和暦入力
            wstr = wstr & "Format(残高" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "融資残高,"
            wstr = wstr & "Format(未払利息増" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "合計未払利息増,"
            wstr = wstr & "Format(未払利息減" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "合計未払利息減,"
            wstr = wstr & "Format(未払利息" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "合計未払利息残高,"
            wstr = wstr & "Format(損益利息額" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "合計損益利息額,"
        Else
        '西暦入力
            wstr = wstr & "Format(残高" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "融資残高,"
            wstr = wstr & "Format(未払利息増" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "合計未払利息増,"
            wstr = wstr & "Format(未払利息減" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "合計未払利息減,"
            wstr = wstr & "Format(未払利息" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "合計未払利息残高,"
            wstr = wstr & "Format(損益利息額" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "合計損益利息額,"
        End If
    Next j
    
    For j = wML To wML
        ws01 = "_" & Right("00" & CStr(j), 2)
        
        If G基本情報.日付入力区分 = "0" Then
        '和暦入力
            wstr = wstr & "Format(残高" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "融資残高,"
            wstr = wstr & "Format(未払利息増" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "合計未払利息増,"
            wstr = wstr & "Format(未払利息減" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "合計未払利息減,"
            wstr = wstr & "Format(未払利息" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "合計未払利息残高,"
            wstr = wstr & "Format(損益利息額" & ws01 & ",'#,##0') As " & w推移表タイトル.X番目年月(j) & "合計損益利息額"
        Else
        '西暦入力
            wstr = wstr & "Format(残高" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "融資残高,"
            wstr = wstr & "Format(未払利息増" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "合計未払利息増,"
            wstr = wstr & "Format(未払利息減" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "合計未払利息減,"
            wstr = wstr & "Format(未払利息" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "合計未払利息残高,"
            wstr = wstr & "Format(損益利息額" & ws01 & ",'#,##0') As " & Format(w推移表タイトル.X番目年月(j), "yyyymm") & "合計損益利息額"
        End If
    Next j

    wstr = wstr & " INTO [Text;database=" & wCsvDir & "]." & "[" & pCsvFileName & "]"
    
    wstr = wstr + " FROM ((((((DCDA010_借入残高推移表結果 As Z"
    wstr = wstr + " Inner Join DCDA010_借入残高推移表結果２ As Z2"
    wstr = wstr + " ON Z.借入番号 = Z2.借入番号)"
    wstr = wstr + " Inner Join " & wsTbl & " As K"
    wstr = wstr + " ON Z.借入番号 = K.借入番号)"
    wstr = wstr + " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr + " ON K.銀行番号 = G.銀行番号)"
    wstr = wstr + " Inner Join DAAA116_借入金種別 As KS"
    wstr = wstr + " ON K.借入金種別区分 = KS.借入金種別区分)"
    wstr = wstr + " LEFT JOIN DAAA200_部門マスタ AS B"
    wstr = wstr + "  ON B.部門番号 = K.プロジェクト番号)"
    wstr = wstr + " Left Join DAAA116_基準金利 As KK"
    wstr = wstr + " ON K.基準金利区分 = KK.基準金利区分)"
    wstr = wstr + " Left Join DAAA115_金利シミュレーショングループ As KG"
    wstr = wstr + " ON K.金利グループ区分 = KG.金利グループ区分"
    
    wstr = wstr & " Where K.利息区分='" & XMXA020_区分("利息区分", "利息後払") & "'"
    wstr = wstr & "  And (未払利息増合計<>0"
    wstr = wstr & " Or 未払利息減合計<>0"
    wstr = wstr & " Or 未払利息合計<>0"
    wstr = wstr & " Or 損益利息額合計<>0"
     
    For j = 1 To wML
        ws01 = "_" & Right("00" & CStr(j), 2)
        
        wstr = wstr & " Or 未払利息増" & ws01 & "<>0"
        wstr = wstr & " Or 未払利息減" & ws01 & "<>0"
        wstr = wstr & " Or 未払利息" & ws01 & "<>0"
        wstr = wstr & " Or 損益利息額" & ws01 & "<>0"
    Next j

    wstr = wstr & " )"
    wstr = wstr & " Order by K.銀行番号,K.金利種別,K.有担保フラグ,Z.借入番号"
    
    GDb.Execute wstr
    
    DoEvents
    
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
End Sub

'------------------------------------------------
' MX040_借入金台帳
'------------------------------------------------
Public Sub MX040_借入金台帳(pCsvFileName As String)
'
    Dim ws01 As String
    Dim wsNendo As String
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wstr = "SELECT "
    wstr = wstr & "  K.借入番号 AS 借入番号"
    wstr = wstr & ", KS.借入金種別名 AS 借入種別" '借入金種別名
    wstr = wstr & ", K.借入内容 AS 借入内容"
    wstr = wstr + ", IIF(K.手入力区分 = " & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",'標準登録','入力登録') As 登録方法"
    'wstr = wstr + ", IIF(K.手入力区分 = " & P8.FCDbl(XMXA020_区分("日割計算区分", "自動計算")) & ",'自動計算','入力登録') As 日割計算"
    
    '借入内容
    wstr = wstr & ", K.銀行番号 AS 銀行番号"
    wstr = wstr & ", G.銀行名 AS 銀行名"
    wstr = wstr & ", FORMAT(K.実行日,'" & Gfmtcsv年月日 & "') AS 実行日"
    wstr = wstr & ", FORMAT(K.初回返済年月,'" & Gfmtcsv年月 & "') AS 初回返済年月"
    wstr = wstr & ", FORMAT(K.初回返済実行日,'" & Gfmtcsv年月日 & "') AS 初回返済実行日"
'    wstr = wstr & ", IIF(K.手入力区分 = " & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",FORMAT(K.金利初回年月,'" & Gfmtcsv年月 & "'),'**') As 金利初回年月"
    wstr = wstr & ", FORMAT(K.金利初回年月,'" & Gfmtcsv年月 & "') As 金利初回年月"
    wstr = wstr & ", FORMAT(K.最終返済年月,'" & Gfmtcsv年月 & "') AS 最終返済年月"
    wstr = wstr & ", FORMAT(K.最終返済実行日,'" & Gfmtcsv年月日 & "') AS 最終返済実行日"
    
    wstr = wstr & ", K.融資金額 AS 融資金額"
'    wstr = wstr & ", IIF(K.手入力区分 = " & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.初回返済額,'**') As 初回返済額"
'    wstr = wstr & ", IIF(K.手入力区分 = " & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.毎月返済額,'**') As 毎月返済額"
'    wstr = wstr & ", IIF(K.手入力区分 = " & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.最終返済額,'**') As 最終返済額"
    wstr = wstr & ", K.初回返済額 As 初回返済額"
    wstr = wstr & ", K.毎月返済額 As 毎月返済額"
    wstr = wstr & ", K.最終返済額 As 最終返済額"
    
    wstr = wstr & ", IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期借入金','長期借入金') AS 長短区分"
    wstr = wstr + ", IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As 金利種別"
    wstr = wstr & ", KK.基準金利名 AS 基準金利名"
    wstr = wstr & ", FORMAT(K.利率,'0.00000') AS 利率"
    wstr = wstr & ", K.金利条件 AS 金利備考" '金利条件
    wstr = wstr & ", IIF(K.設備フラグ=0,'運転資金','設備資金') AS 資金区分" '設備フラグ
    wstr = wstr & ", K.資金用途 AS 資金用途"
    wstr = wstr & ", FORMAT(K.解約実行日,'" & Gfmtcsv年月日 & "') AS 解約実行日"
    wstr = wstr + ", IIF(K.有担保フラグ=0,'無担保','有担保') As 担保区分"
    wstr = wstr & ", K.担保名 AS 担保名"
    
    '支払内容
    wstr = wstr & ", IIF(K.支払日=31,'月末',K.支払日) AS 支払日"
    'wstr = wstr & ", IIF(K.手入力区分 = " & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.返済単位月数,'**') As 返済単位" '返済単位月数
    wstr = wstr & ", K.返済単位月数 As 返済単位" '返済単位月数
    wstr = wstr + ", IIF(K.営業日区分 = " & P8.FCDbl(XMXA020_区分("営業日", "翌営業日")) & ",'翌営業日','前営業日') As 営業日"
    wstr = wstr + ", IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As 利息区分"
    wstr = wstr + ", IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "控除無し")) & ",'控除無し',"
    wstr = wstr + "  IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "実行日控除")) & ",'実行日控除',"
    wstr = wstr + "  IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "最終返済日控除")) & ",'最終返済日控除','実行日及び最終返済日控除'))) As 利息控除区分"
    wstr = wstr + ", IIF(K.利息計算日数区分 = " & P8.FCDbl(XMXA020_区分("利息計算日数", "営業日数")) & ",'営業日数','固定日数') As 利息計算日数"
    wstr = wstr + ", IIF(K.利息支払方法 = " & P8.FCDbl(XMXA020_区分("利息支払", "毎月")) & ",'毎月','一括') As 利息支払方法"
    wstr = wstr + ", IIF(K.金利計算年間日数 = " & P8.FCDbl(XMXA020_区分("金利計算", "365日")) & ",'365日','360日') As 金利計算日数" '金利計算年間日数
    
    'シミュレーション
    wstr = wstr & ", K.金融リストラ番号 AS 金融リストラ番号"
    wstr = wstr & ", IIF(K.sm区分=0,'OFF','ON') As SM区分"
    wstr = wstr & ", FORMAT(K.金融解約実行日,'" & Gfmtcsv年月日 & "') AS 金融解約実行日"
    wstr = wstr & ", KSM.金利グループ名 AS 金利グループ名"
    
    '口座情報
    wstr = wstr & ", G.金融機関番号 AS 金融機関番号"
    wstr = wstr & ", G.金融機関名 AS 金融機関名"
    wstr = wstr & ", G.支店番号 AS 支店番号"
    wstr = wstr & ", G.支店名 AS 支店名"
    wstr = wstr & ", G.預金種別 AS 預金種別"
    wstr = wstr & ", G.口座番号 AS 口座番号"
    
    wstr = wstr & ", B.部門番号 AS 部門番号"
    wstr = wstr & ", B.部門名 AS 部門名"
    wstr = wstr & ", B.部門略名 AS 部門略名"
    
    wstr = wstr & " FROM ((((DBDA010_借入金 AS K LEFT JOIN DAAA040_銀行マスタ AS G"
    wstr = wstr & "  ON G.銀行番号 = K.銀行番号)"
    wstr = wstr & " LEFT JOIN DAAA116_借入金種別 AS KS"
    wstr = wstr & "  ON KS.借入金種別区分 = K.借入金種別区分)"
    wstr = wstr & " LEFT JOIN DAAA200_部門マスタ AS B"
    wstr = wstr & "  ON B.部門番号 = K.プロジェクト番号)"
    wstr = wstr & " LEFT JOIN DAAA115_金利シミュレーショングループ AS KSM"
    wstr = wstr & "  ON KSM.金利グループ区分 = K.金利グループ区分)"
    wstr = wstr & " LEFT JOIN DAAA116_基準金利 AS KK"
    wstr = wstr & "  ON KK.基準金利区分 = K.基準金利区分"
    wstr = wstr & " WHERE K.借入番号='" & GRpt.コンボ_01 & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
        
    On Error GoTo Err_Hundle
        '名称
        Write #1, _
            "借入番号", "借入種別", "借入内容", "登録方法", _
            "銀行番号", "銀行名", _
            "実行日", "初回返済年月", "初回返済実行日", "金利初回年月", "最終返済年月", "最終返済実行日", _
            "融資金額", "初回返済額", "毎月返済額", "最終返済額", _
            "長短区分", "金利種別", "基準金利名", "利率", "金利備考", _
            "資金区分", "資金用途", _
            "解約実行日", _
            "担保区分", "担保名", _
            "支払日", "返済単位", _
            "営業日", "利息区分", "利息控除区分", "利息計算日数", "利息支払方法", "金利計算日数", _
            "金融リストラ番号", "SM区分", "金融解約実行日", "金利グループ名", _
            "金融機関番号", "金融機関名", _
            "支店番号", "支店名", _
            "預金種別", "口座番号", _
            "部門番号", "部門名", "部門略名"
    
        Do Until wRs.eof
            Write #1, _
            P8.FCStr(wRs.Fields("借入番号").Value), P8.FCStr(wRs.Fields("借入種別").Value), P8.FCStr(wRs.Fields("借入内容").Value), _
            P8.FCStr(wRs.Fields("登録方法").Value), _
            P8.FCStr(wRs.Fields("銀行番号").Value), P8.FCStr(wRs.Fields("銀行名").Value), _
            P8.FCStr(wRs.Fields("実行日").Value), _
            P8.FCStr(wRs.Fields("初回返済年月").Value), P8.FCStr(wRs.Fields("初回返済実行日").Value), _
            P8.FCStr(wRs.Fields("金利初回年月").Value), _
            P8.FCStr(wRs.Fields("最終返済年月").Value), P8.FCStr(wRs.Fields("最終返済実行日").Value), _
            P8.FCDbl(wRs.Fields("融資金額").Value), P8.FCDbl(wRs.Fields("初回返済額").Value), P8.FCDbl(wRs.Fields("毎月返済額").Value), P8.FCDbl(wRs.Fields("最終返済額").Value), _
            P8.FCStr(wRs.Fields("長短区分").Value), P8.FCStr(wRs.Fields("金利種別").Value), P8.FCStr(wRs.Fields("基準金利名").Value), _
            P8.FCDbl(wRs.Fields("利率").Value), _
            P8.FCStr(wRs.Fields("金利備考").Value), _
            P8.FCStr(wRs.Fields("資金区分").Value), P8.FCStr(wRs.Fields("資金用途").Value), _
            P8.FCStr(wRs.Fields("解約実行日").Value), _
            P8.FCStr(wRs.Fields("担保区分").Value), P8.FCStr(wRs.Fields("担保名").Value), _
            P8.FCStr(wRs.Fields("支払日").Value), P8.FCStr(wRs.Fields("返済単位").Value), _
            P8.FCStr(wRs.Fields("営業日").Value), P8.FCStr(wRs.Fields("利息区分").Value), _
            P8.FCStr(wRs.Fields("利息控除区分").Value), P8.FCStr(wRs.Fields("利息計算日数").Value), _
            P8.FCStr(wRs.Fields("利息支払方法").Value), P8.FCStr(wRs.Fields("金利計算日数").Value), _
            P8.FCStr(wRs.Fields("金融リストラ番号").Value), P8.FCStr(wRs.Fields("SM区分").Value), P8.FCStr(wRs.Fields("金融解約実行日").Value), _
            P8.FCStr(wRs.Fields("金利グループ名").Value), _
            P8.FCStr(wRs.Fields("金融機関番号").Value), P8.FCStr(wRs.Fields("金融機関名").Value), _
            P8.FCStr(wRs.Fields("支店番号").Value), P8.FCStr(wRs.Fields("支店名").Value), _
            P8.FCStr(wRs.Fields("預金種別").Value), P8.FCStr(wRs.Fields("口座番号").Value), _
            P8.FCStr(wRs.Fields("部門番号").Value), P8.FCStr(wRs.Fields("部門名").Value), P8.FCStr(wRs.Fields("部門略名").Value)
            
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
'
'    GDb.Execute wstr
'
'    DoEvents
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_借入明細表
'------------------------------------------------
Public Sub MX040_借入明細表(pCsvFileName As String)
'
    Dim ws01 As String
    Dim wsNendo As String
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wstr = "SELECT "
    If GRpt.帳票名 = "借入明細表" Then
        wstr = wstr & "M.借入番号,"
    ElseIf GRpt.帳票名 = "貸付明細表" Then
        wstr = wstr & "M.借入番号 As 貸付番号,"
    End If
    
    '2016/03/29 利子補給に伴う変更
    wstr = wstr & "KS.利子補給金フラグ,"
    
    wstr = wstr & "M.返済回数,"
    'wstr = wstr & "Format(M.実際年月日,'yyyy/mm/dd') As 返済年月日,"
    'wstr = wstr & "Format(M.利息計算年月日,'yyyy/mm/dd') As 利息計算年月日,"
    '2012.10.23　追加 by k.kunita
    wstr = wstr & "Format(M.実際年月日,'" & Gfmtcsv年月日 & "') As 返済年月日,"
    wstr = wstr & "Format(M.利息計算年月日,'" & Gfmtcsv年月日 & "') As 利息計算年月日,"
    wstr = wstr & "M.元金額,"
    wstr = wstr & "M.利息額,"
    wstr = wstr & "M.仮計上利息額 As 調整利息額,"
    wstr = wstr & "M.返済金額,"
    wstr = wstr & "M.融資残高,"
    wstr = wstr & "M.日割日数,"
    wstr = wstr & "M.利息対象期間日数 As 利息対象期間日数,"
    'wstr = wstr & "M.保証料 As 保証料,"
    'wstr = wstr & "M.金融保証料 As 金融保証料,"
    wstr = wstr & "format(M.利率,'#,##0.00000') As 利率"
    'wstr = wstr & " INTO [Text;database=" & wCsvDir & "]." & "[" & pCsvFileName & "]"
    wstr = wstr & " FROM (DCDA020_借入金明細 As M"
    wstr = wstr & " INNER JOIN DBDA010_借入金 As K"
    wstr = wstr & " ON M.借入番号 = K.借入番号)"
    wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As KS"
    wstr = wstr & " ON K.借入金種別区分 = KS.借入金種別区分"
    wstr = wstr & " WHERE M.借入番号='" & GRpt.コンボ_01 & "'"
    wstr = wstr & " Order by M.実際年月日"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
        
    '名称
    '16/03/26 利子補給に伴う変更
    Write #1, _
        CStr(wRs.Fields("借入番号").Name), _
        CStr(wRs.Fields("返済回数").Name), _
        CStr(wRs.Fields("返済年月日").Name), _
        CStr(wRs.Fields("利息計算年月日").Name), _
        CStr(wRs.Fields("元金額").Name), _
        IIf(wRs.Fields("利子補給金フラグ").Value = 0, CStr(wRs.Fields("利息額").Name), "利子補給金"), _
        CStr(wRs.Fields("調整利息額").Name), _
        CStr(wRs.Fields("返済金額").Name), _
        CStr(wRs.Fields("融資残高").Name), _
        CStr(wRs.Fields("日割日数").Name), _
        CStr(wRs.Fields("利息対象期間日数").Name), _
        CStr(wRs.Fields("利率").Name)
    
    On Error GoTo Err_Hundle
    
        Do Until wRs.eof
        
            Write #1, _
                P8.FCStr(wRs.Fields("借入番号").Value), _
                P8.FCStr(wRs.Fields("返済回数").Value), _
                P8.FCStr(wRs.Fields("返済年月日").Value), _
                P8.FCStr(wRs.Fields("利息計算年月日").Value), _
                P8.FCDbl(wRs.Fields("元金額").Value), _
                P8.FCDbl(wRs.Fields("利息額").Value), _
                P8.FCDbl(wRs.Fields("調整利息額").Value), _
                P8.FCDbl(wRs.Fields("返済金額").Value), _
                P8.FCDbl(wRs.Fields("融資残高").Value), _
                P8.FCDbl(wRs.Fields("日割日数").Value), _
                P8.FCDbl(wRs.Fields("利息対象期間日数").Value), _
                P8.FCDbl(wRs.Fields("利率").Value)
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
'
'    GDb.Execute wstr
'
'    DoEvents
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_社債明細表
'------------------------------------------------
Public Sub MX040_社債明細表(pCsvFileName As String)
'
    Dim ws01 As String
    Dim wsNendo As String
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wstr = "SELECT "
    wstr = wstr & "M.借入番号,"
    wstr = wstr & "M.返済回数,"
    wstr = wstr & "Format(M.実際年月日,'yyyy/mm/dd') As 返済年月日,"
    wstr = wstr & "Format(M.利息計算年月日,'yyyy/mm/dd') As 利息計算年月日,"
    wstr = wstr & "M.元金額,"
    wstr = wstr & "M.利息額,"
    wstr = wstr & "M.仮計上利息額 As 調整利息額,"
    wstr = wstr & "M.返済金額,"
    wstr = wstr & "M.融資残高,"
    wstr = wstr & "M.日割日数,"
    wstr = wstr & "M.利息対象期間日数 As 利息対象期間日数,"
    'wstr = wstr & "M.保証料 As 保証料,"
    'wstr = wstr & "M.金融保証料 As 金融保証料,"
    wstr = wstr & "format(M.利率,'#,##0.00000') As 利率,"
    wstr = wstr & "M.初期手数料,"
    wstr = wstr & "M.元金手数料,"
    wstr = wstr & "M.利息手数料,"
    wstr = wstr & "M.初期手数料+M.元金手数料+M.利息手数料 As 手数料計,"
    wstr = wstr & "M.保証料,"
    wstr = wstr & "M.元金額+M.利息額+M.初期手数料+M.元金手数料+M.利息手数料+M.保証料 As 支払計"
    'wstr = wstr & " INTO [Text;database=" & wCsvDir & "]." & "[" & pCsvFileName & "]"
    wstr = wstr & " FROM DCDA020_借入金明細 As M"
    wstr = wstr & " INNER JOIN DBDA010_借入金 As K"
    wstr = wstr & " ON M.借入番号 = K.借入番号"
    wstr = wstr & " WHERE M.借入番号='" & GRpt.コンボ_01 & "'"
    wstr = wstr & " Order by M.実際年月日"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        '名称
        Write #1, _
            CStr(wRs.Fields("借入番号").Name), _
            CStr(wRs.Fields("返済回数").Name), _
            CStr(wRs.Fields("返済年月日").Name), _
            CStr(wRs.Fields("利息計算年月日").Name), _
            CStr(wRs.Fields("元金額").Name), _
            CStr(wRs.Fields("利息額").Name), _
            CStr(wRs.Fields("調整利息額").Name), _
            CStr(wRs.Fields("返済金額").Name), _
            CStr(wRs.Fields("融資残高").Name), _
            CStr(wRs.Fields("日割日数").Name), _
            CStr(wRs.Fields("利息対象期間日数").Name), _
            CStr(wRs.Fields("利率").Name), _
            CStr(wRs.Fields("初期手数料").Name), _
            CStr(wRs.Fields("元金手数料").Name), _
            CStr(wRs.Fields("利息手数料").Name), _
            CStr(wRs.Fields("手数料計").Name), _
            CStr(wRs.Fields("保証料").Name), _
            CStr(wRs.Fields("支払計").Name)
        
        Do Until wRs.eof
            Write #1, _
                P8.FCStr(wRs.Fields("借入番号").Value), _
                P8.FCStr(wRs.Fields("返済回数").Value), _
                P8.FCStr(wRs.Fields("返済年月日").Value), _
                P8.FCStr(wRs.Fields("利息計算年月日").Value), _
                P8.FCDbl(wRs.Fields("元金額").Value), _
                P8.FCDbl(wRs.Fields("利息額").Value), _
                P8.FCDbl(wRs.Fields("調整利息額").Value), _
                P8.FCDbl(wRs.Fields("返済金額").Value), _
                P8.FCDbl(wRs.Fields("融資残高").Value), _
                P8.FCDbl(wRs.Fields("日割日数").Value), _
                P8.FCDbl(wRs.Fields("利息対象期間日数").Value), _
                P8.FCDbl(wRs.Fields("利率").Value), _
                P8.FCDbl(wRs.Fields("初期手数料").Value), _
                P8.FCDbl(wRs.Fields("元金手数料").Value), _
                P8.FCDbl(wRs.Fields("利息手数料").Value), _
                P8.FCDbl(wRs.Fields("手数料計").Value), _
                P8.FCDbl(wRs.Fields("保証料").Value), _
                P8.FCDbl(wRs.Fields("支払計").Value)
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
'
'    GDb.Execute wstr
'
'    DoEvents
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_利息前払未払明細表
'------------------------------------------------
Public Sub MX040_利息前払未払明細表(pCsvFileName As String)
'
    Dim ws01 As String, ws02 As String
    Dim wd利息計算対象額1 As Double, wd利息計算対象額2 As Double
    Dim wd利息額1 As Double, wd利息額2 As Double
    Dim wi日数1 As Integer, wi日数2 As Integer
    Dim ws利率1 As Single, ws利率2 As Single
    'Dim ws番号1 As String, ws番号2 As String
    Dim ws年月日1 As String, ws年月日2 As String
    Dim ws期間1 As String, ws期間2 As String
    Dim ws式1 As String, ws式2 As String
    Dim wd01 As Double
    Dim wd利息増 As Double, wd利息減 As Double, wd利息残高 As Double
    
    '2014/08/26 計算式修正
    Dim wi利息期間対象日数1 As Integer
    Dim wd利息期間対象額1 As Double
    Dim wi利息調整F1 As Integer
    Dim wd引算1 As Double
    Dim wd前月前払利息残高 As Double
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wd利息増 = 0: wd利息減 = 0: wd利息残高 = 0
    
    wstr = ""
    wstr = wstr & "Select "
    'wstr = wstr & "Format(KM.締年月,'" & Gfmt年月 & "') As 年月,"
    'wstr = wstr & "Format(KM.返済年月日,'" & Gfmt年月日 & "') As 年月日,"
    wstr = wstr & "Format(KM.締年月,'yyyy/mm') As 年月,"
    wstr = wstr & "Format(KM.返済年月日,'yyyy/mm/dd') As 年月日,"
    'wstr = wstr & "KM.月毎NO As 番号,"
    wstr = wstr & "KM.利息計算対象額,"
    wstr = wstr & "KM.利息額増,"
    wstr = wstr & "KM.利息額減,"
    wstr = wstr & "KM.日割日数,"
    wstr = wstr & "Format(KM.利率,'#,##0.00000') As 利率,"
    'wstr = wstr & "Format(KM.開始年月日,'" & Gfmt年月日 & "') As 開始日,"
    'wstr = wstr & "Format(KM.終了年月日,'" & Gfmt年月日 & "') As 終了日,"
    wstr = wstr & "Format(KM.開始年月日,'yyyy/mm/dd') As 開始日,"
    wstr = wstr & "Format(KM.終了年月日,'yyyy/mm/dd') As 終了日,"
    
    '2014/08/26 計算式修正
    wstr = wstr & "KM.利息期間対象額,"
    wstr = wstr & "KM.利息期間対象日数,"
    wstr = wstr & "KM.利息調整F,"
    
    wstr = wstr & "K.借入番号,"
    wstr = wstr & "S.支払区分名 As 支払日,"
    wstr = wstr & "IIF(K.営業日区分 = " & P8.FCDbl(XMXA020_区分("営業日", "翌営業日")) & ",'翌営業日','前営業日') As 営業日,"
    wstr = wstr & "IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As 利息区分,"
    wstr = wstr & "IIF(K.利息計算日数区分 = " & P8.FCDbl(XMXA020_区分("利息計算日数", "営業日数")) & ",'営業日数','固定日数') As 利息日数,"
    wstr = wstr & "IIF(K.利息支払方法 = " & P8.FCDbl(XMXA020_区分("利息支払", "毎月")) & ",'毎月','一括') As 利息支払方法,"
    wstr = wstr & "IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "控除無し")) & ",'控除無し',"
    wstr = wstr & "IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "実行日控除")) & ",'実行日控除',"
    wstr = wstr & "IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "最終返済日控除")) & ",'最終返済日控除',"
    wstr = wstr & "IIF(K.利息控除区分 = " & P8.FCDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) & ",'実行日及び最終返済日控除','中間利払最終日控除')))) As 利息控除区分,"
    wstr = wstr & "IIF(K.金利計算年間日数 = " & P8.FCDbl(XMXA020_区分("金利計算", "365日")) & ",'365','360') As 金利年間日数"
    
    wstr = wstr + " From (DCDA030_利息未払前払明細 As KM"
    wstr = wstr + " Inner Join DBDA010_借入金 As K"
    wstr = wstr + " ON KM.借入番号 = K.借入番号)"
    wstr = wstr + " Inner Join DAAB020_支払区分マスタ As S"
    wstr = wstr + " ON K.支払日 = S.支払日"
    
    '2014/08/26 計算式修正
    'wstr = wstr & " Order BY 返済年月日,利息額増 desc,KM.月毎NO"
    wstr = wstr & " Order BY 返済年月日,利息調整F desc"
    
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        '名称
        If P8.FCStr(wRs.Fields("利息区分").Value) = "利息先払" Then
            Write #1, _
                "年月", _
                "利息計上日", "前払利息減", "期間", "日割日数", "利率", "式", _
                "支払日", "前払利息増", "期間", "日割日数", "利率", "式", _
                "前払利息増", "前払利息減", "前払利息残高", _
                "借入番号", "支払日", "営業日", "利息区分", "利息日数", "利息支払方法", "利息控除区分", "金利年間日数"
        ElseIf P8.FCStr(wRs.Fields("利息区分").Value) = "利息後払" Then
            Write #1, _
                "年月", _
                "支払日", "未払利息増", "期間", "日割日数", "利率", "式", _
                "利息計上日", "未払利息減", "期間", "日割日数", "利率", "式", _
                "未払利息増", "未払利息減", "未払利息残高", _
                "借入番号", "支払日", "営業日", "利息区分", "利息日数", "利息支払方法", "利息控除区分", "金利年間日数"
        End If
    
    Do Until wRs.eof
        
        wd利息計算対象額1 = 0: wd利息計算対象額2 = 0
        wd利息額1 = 0: wd利息額2 = 0
        wi日数1 = 0: wi日数2 = 0
        ws利率1 = 0: ws利率2 = 0
        'ws番号1 = "": ws番号2 = ""
        ws年月日1 = "": ws年月日2 = ""
        ws期間1 = "": ws期間2 = ""
        ws式1 = "": ws式2 = ""
        
        '2014/08/26 計算式修正
        wd利息期間対象額1 = 0
        wi利息期間対象日数1 = 0
        wi利息調整F1 = 0
        
        If P8.FCStr(wRs.Fields("利息区分").Value) = "利息先払" Then
        '利息先払 1：減、2：増
            If P8.FCDbl(wRs.Fields("利息額減").Value) <> 0 Then
                'ws番号1 = P8.FCStr(wRs.Fields("番号").Value)
                ws年月日1 = P8.FCStr(wRs.Fields("年月日").Value)
                wd利息計算対象額1 = P8.FCDbl(wRs.Fields("利息計算対象額").Value)
                wd利息額1 = P8.FCDbl(wRs.Fields("利息額減").Value)
                ws期間1 = P8.FCStr(wRs.Fields("開始日").Value) & "～" & P8.FCStr(wRs.Fields("終了日").Value)
                wi日数1 = P8.FCDbl(wRs.Fields("日割日数").Value)
                ws利率1 = P8.FCDbl(wRs.Fields("利率").Value)
                ws式1 = "= " & Format(wd利息計算対象額1, "#,##0") & " × " & Format(ws利率1, "#,##0.00000") & "%" & " × " & wi日数1 & " / " & P8.FCStr(wRs.Fields("金利年間日数").Value)
            
                '2014/08/26 計算式修正
                wd利息期間対象額1 = P8.FCDbl(wRs.Fields("利息期間対象額").Value)
                wi利息期間対象日数1 = P8.FCDbl(wRs.Fields("利息期間対象日数").Value)
                wi利息調整F1 = P8.FCDbl(wRs.Fields("利息調整F").Value)
                If wi利息調整F1 = 0 Then
                    wd引算1 = wd引算1 + wd利息額1
                    ws式1 = "= " & Format(wd利息期間対象額1, "#,##0") & " × " & wi日数1 & " / " & wi利息期間対象日数1
                
                ElseIf wi利息調整F1 = 2 Then
                '解約 利息先払
                    '前月前払利息残高＝－前払利息増＋前払利息減
                    '前払利息減＝前月前払利息残高＋前払利息増
                    wd前月前払利息残高 = -wd利息期間対象額1 + wd利息額1
                    ws式1 = "=" & Format(wd前月前払利息残高, "#,##0") & " + （" & Format(wd利息期間対象額1, "#,##0") & "）"
                    wd引算1 = 0
                Else
                    If wd引算1 = 0 Then
                        ws式1 = "= " & Format(wd利息期間対象額1, "#,##0") & " × " & wi日数1 & " / " & wi利息期間対象日数1
                    Else
                        ws式1 = "= " & Format(wd利息期間対象額1, "#,##0") & " － " & Format(wd引算1, "#,##0")
                        wd引算1 = 0
                    End If
                End If
                
            End If
            
            If P8.FCDbl(wRs.Fields("利息額増").Value) <> 0 Then
                'ws番号2 = P8.FCStr(wRs.Fields("番号").Value)
                ws年月日2 = P8.FCStr(wRs.Fields("年月日").Value)
                wd利息計算対象額2 = P8.FCDbl(wRs.Fields("利息計算対象額").Value)
                wd利息額2 = P8.FCDbl(wRs.Fields("利息額増").Value)
                ws期間2 = P8.FCStr(wRs.Fields("開始日").Value) & "～" & P8.FCStr(wRs.Fields("終了日").Value)
                wi日数2 = P8.FCDbl(wRs.Fields("日割日数").Value)
                ws利率2 = P8.FCDbl(wRs.Fields("利率").Value)
                ws式2 = "= " & Format(wd利息計算対象額2, "#,##0") & " × " & Format(ws利率2, "#,##0.00000") & "%" & " × " & wi日数2 & " / " & P8.FCStr(wRs.Fields("金利年間日数").Value)
            End If
        End If
'
        If P8.FCStr(wRs.Fields("利息区分").Value) = "利息後払" Then
        '利息後払 1：増、2：減
            If P8.FCDbl(wRs.Fields("利息額増").Value) <> 0 Then
                'ws番号1 = P8.FCStr(wRs.Fields("番号").Value)
                ws年月日1 = P8.FCStr(wRs.Fields("年月日").Value)
                wd利息計算対象額1 = P8.FCDbl(wRs.Fields("利息計算対象額").Value)
                wd利息額1 = P8.FCDbl(wRs.Fields("利息額増").Value)
                ws期間1 = P8.FCStr(wRs.Fields("開始日").Value) & "～" & P8.FCStr(wRs.Fields("終了日").Value)
                wi日数1 = P8.FCDbl(wRs.Fields("日割日数").Value)
                ws利率1 = P8.FCDbl(wRs.Fields("利率").Value)
                ws式1 = "= " & Format(wd利息計算対象額1, "#,##0") & " × " & Format(ws利率1, "#,##0.00000") & "%" & " × " & wi日数1 & " / " & P8.FCStr(wRs.Fields("金利年間日数").Value)
            
                '2014/08/26 計算式修正
                wd利息期間対象額1 = P8.FCDbl(wRs.Fields("利息期間対象額").Value)
                wi利息期間対象日数1 = P8.FCDbl(wRs.Fields("利息期間対象日数").Value)
                wi利息調整F1 = P8.FCDbl(wRs.Fields("利息調整F").Value)
                If wi利息調整F1 = 0 Then
                    wd引算1 = wd引算1 + wd利息額1
                    ws式1 = "= " & Format(wd利息期間対象額1, "#,##0") & " × " & wi日数1 & " / " & wi利息期間対象日数1
                Else
                    If wd引算1 = 0 Then
                        ws式1 = "= " & Format(wd利息期間対象額1, "#,##0") & " × " & wi日数1 & " / " & wi利息期間対象日数1
                    Else
                        ws式1 = "= " & Format(wd利息期間対象額1, "#,##0") & " － " & Format(wd引算1, "#,##0")
                        wd引算1 = 0
                    End If
                End If
                
            End If
            
            If P8.FCDbl(wRs.Fields("利息額減").Value) <> 0 Then
                'ws番号2 = P8.FCStr(wRs.Fields("番号").Value)
                ws年月日2 = P8.FCStr(wRs.Fields("年月日").Value)
                wd利息計算対象額2 = P8.FCDbl(wRs.Fields("利息計算対象額").Value)
                wd利息額2 = P8.FCDbl(wRs.Fields("利息額減").Value)
                ws期間2 = P8.FCStr(wRs.Fields("開始日").Value) & "～" & P8.FCStr(wRs.Fields("終了日").Value)
                wi日数2 = P8.FCDbl(wRs.Fields("日割日数").Value)
                ws利率2 = P8.FCDbl(wRs.Fields("利率").Value)
                ws式2 = "= " & Format(wd利息計算対象額2, "#,##0") & " × " & Format(ws利率2, "#,##0.00000") & "%" & " × " & wi日数2 & " / " & P8.FCStr(wRs.Fields("金利年間日数").Value)
            End If
        End If
            
        wd01 = wd利息残高 + P8.FCDbl(wRs.Fields("利息額増").Value) - P8.FCDbl(wRs.Fields("利息額減").Value)
        
        Write #1, _
            P8.FCStr(wRs.Fields("年月").Value), _
            ws年月日1, wd利息額1, ws期間1, wi日数1, ws利率1, ws式1, _
            ws年月日2, wd利息額2, ws期間2, wi日数2, ws利率2, ws式2, _
            P8.FCDbl(wRs.Fields("利息額増").Value), P8.FCDbl(wRs.Fields("利息額減").Value), _
            wd01, _
            P8.FCStr(wRs.Fields("借入番号").Value), _
            P8.FCStr(wRs.Fields("支払日").Value), P8.FCStr(wRs.Fields("営業日").Value), _
            P8.FCStr(wRs.Fields("利息区分").Value), P8.FCStr(wRs.Fields("利息日数").Value), _
            P8.FCStr(wRs.Fields("利息支払方法").Value), P8.FCStr(wRs.Fields("利息控除区分").Value), _
            P8.FCStr(wRs.Fields("金利年間日数").Value)
            
        wd利息残高 = wd01

    wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_借入一覧表
'------------------------------------------------
Private Sub MX040_借入一覧表(pCsvFileName As String)
'
    Dim ws01 As String
    Dim wsNendo As String, wsTbl As String
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wsTbl = "DBDA010_借入金"
    If GRpt.帳票名 <> "借入一覧表" Then
        wsTbl = "DBDA010_貸付金"
    End If
'
    wstr = "SELECT "
    
    If GRpt.帳票名 = "借入一覧表" Then
        wstr = wstr & "K.借入番号,"
        wstr = wstr & "K.借入内容,"
    Else
        wstr = wstr & "K.借入番号 As 貸付番号,"
        wstr = wstr & "K.借入内容 As 貸付内容,"
    End If
    
    wstr = wstr & " KS.借入金種別名,"
    wstr = wstr & " K.銀行番号,"
    wstr = wstr & " G.銀行名,"
    wstr = wstr + "B.部門名,"
    
    wstr = wstr & "IIF(K.営業日区分 = " & P8.FCDbl(XMXA020_区分("営業日", "翌営業日")) & ",'翌営業日','前営業日') As 営業日,"
    wstr = wstr & "IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As 利息区分,"
    wstr = wstr & "IIF(K.利息計算日数区分 = " & P8.FCDbl(XMXA020_区分("利息計算日数", "営業日数")) & ",'営業日数','固定日数') As 利息計算日数,"
    wstr = wstr & "IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As 金利種別,"
    wstr = wstr & "KK.基準金利名,"
    wstr = wstr & "K.金利条件 As 金利備考,"
    
    wstr = wstr & "IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期借入金','長期借入金') AS 長短区分,"
    wstr = wstr & "IIF(K.有担保フラグ = " & P8.FCDbl(XMXA020_区分("有担フラグ", "無担保")) & ",'無担保','有担保') As 担保区分,"
    wstr = wstr & "K.担保名 As 担保内容,"
    wstr = wstr & "IIF(設備フラグ=1,'設備','運転資金') As 設備区分,"
    wstr = wstr & "K.資金用途,"
    
    wstr = wstr & " KG.金利グループ名,"
    wstr = wstr & " IIF(K.sm区分=1 and K.金融リストラ番号<>'','借入SM',IIF(K.sm区分=0 and not (K.金融解約実行日 is null),'解約SM','')) As ｼﾐｭﾚｰｼｮﾝ内容,"
    wstr = wstr & " K.金融リストラ番号 As ｼﾐｭﾚｰｼｮﾝ番号,"
'    wstr = wstr & " format(K.金融解約実行日,'yyyy/mm/dd') As ｼﾐｭﾚｰｼｮﾝ解約日,"
    
    'wstr = wstr & "Format(K.保証料率,'#,##0.00000') As 保証料率,"
    
'    wstr = wstr & "Format(KI.実行日,'yyyy/mm/dd') As 実行日,"
'    wstr = wstr & "Format(KI.初回返済年月,'yyyy/mm/dd') As 初回返済年月,"
'    wstr = wstr & "Format(K.初回返済実行日,'yyyy/mm/dd') As 初回返済年月日,"
'    wstr = wstr & "Format(KI.最終返済年月,'yyyy/mm/dd') As 最終返済年月,"
'    wstr = wstr & "Format(KI.最終返済実行日,'yyyy/mm/dd') As 最終返済年月日,"
'    wstr = wstr & "Format(KI.解約実行日,'yyyy/mm/dd') As 解約年月日,"
    
    '2012.10.24　追加 by k.kuni
    wstr = wstr & " format(K.金融解約実行日,'" & Gfmtcsv年月日 & "') As ｼﾐｭﾚｰｼｮﾝ解約日,"
    
    wstr = wstr & "Format(KI.実行日,'" & Gfmtcsv年月日 & "') As 実行日,"
    wstr = wstr & "Format(KI.初回返済年月,'" & Gfmtcsv年月日 & "') As 初回返済年月,"
    wstr = wstr & "Format(K.初回返済実行日,'" & Gfmtcsv年月日 & "') As 初回返済年月日,"
    wstr = wstr & "Format(KI.最終返済年月,'" & Gfmtcsv年月日 & "') As 最終返済年月,"
    wstr = wstr & "Format(KI.最終返済実行日,'" & Gfmtcsv年月日 & "') As 最終返済年月日,"
    wstr = wstr & "Format(KI.解約実行日,'" & Gfmtcsv年月日 & "') As 解約年月日,"
    wstr = wstr & "IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",KI.初回返済額,0) As 初回返済額,"
    wstr = wstr & "IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",KI.毎月返済額,0) As 毎月返済額,"
    wstr = wstr & "IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",KI.最終返済額,0) As 最終返済額,"
    
    wstr = wstr + "IIF(KI.変動利率フラグ=1,'*','') As 変動利率フラグ,"
    wstr = wstr & "format(KI.利率,'#,##0.00000') As 利率,"
    wstr = wstr & "KI.返済単位月数,"
    
    wstr = wstr & "KI.融資金額 As 融資金額,"
    'wstr = wstr & "Format(KI.入力残高年月,'yyyy/mm/dd') As 残高年月,"
    wstr = wstr & "KI.融資残高 As 融資残高"
    
    wstr = wstr & " From (((((DCDA030_借入一覧表 As KI"
    wstr = wstr + " Inner Join " & wsTbl & " As K"
    wstr = wstr + " ON KI.借入番号 = K.借入番号)"
    wstr = wstr + " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr + " ON K.銀行番号 = G.銀行番号)"
    wstr = wstr + " Inner Join DAAA116_借入金種別 As KS"
    wstr = wstr + " ON K.借入金種別区分 = KS.借入金種別区分)"
    wstr = wstr + " LEFT JOIN DAAA200_部門マスタ AS B"
    wstr = wstr + "  ON B.部門番号 = K.プロジェクト番号)"
    wstr = wstr + " Left Join DAAA116_基準金利 As KK"
    wstr = wstr + " ON K.基準金利区分 = KK.基準金利区分)"
    wstr = wstr + " Left Join DAAA115_金利シミュレーショングループ As KG"
    wstr = wstr + " ON K.金利グループ区分 = KG.金利グループ区分"
    
    wstr = wstr + " Order BY KI.銀行番号,KI.借入番号"
    
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        '名称
        Write #1, _
            "借入番号", "借入内容", "借入金種別名", "銀行番号", "銀行名", "部門名", _
            "営業日", "利息区分", "利息計算日数", "金利種別", "基準金利名", "金利備考", _
            "長短区分", "担保区分", "担保内容", "設備区分", "資金用途", _
            "金利グループ名", "ｼﾐｭﾚｰｼｮﾝ内容", "ｼﾐｭﾚｰｼｮﾝ番号", "ｼﾐｭﾚｰｼｮﾝ解約日", _
            "実行日", "初回返済年月", "初回返済年月日", "最終返済年月", "最終返済年月日", "解約年月日", _
            "初回返済額", "毎月返済額", "最終返済額", _
            "変動利率フラグ", "利率", _
            "返済単位月数", _
            "融資金額", _
            "融資残高"
            
    Do Until wRs.eof
        Write #1, _
            P8.FCStr(wRs.Fields("借入番号").Value), P8.FCStr(wRs.Fields("借入内容").Value), _
            P8.FCStr(wRs.Fields("借入金種別名").Value), _
            P8.FCStr(wRs.Fields("銀行番号").Value), P8.FCStr(wRs.Fields("銀行名").Value), P8.FCStr(wRs.Fields("部門名").Value), _
            P8.FCStr(wRs.Fields("営業日").Value), P8.FCStr(wRs.Fields("利息区分").Value), P8.FCStr(wRs.Fields("利息計算日数").Value), _
            P8.FCStr(wRs.Fields("金利種別").Value), P8.FCStr(wRs.Fields("基準金利名").Value), P8.FCStr(wRs.Fields("金利備考").Value), _
            P8.FCStr(wRs.Fields("長短区分").Value), P8.FCStr(wRs.Fields("担保区分").Value), P8.FCStr(wRs.Fields("担保内容").Value), _
            P8.FCStr(wRs.Fields("設備区分").Value), P8.FCStr(wRs.Fields("資金用途").Value), _
            P8.FCStr(wRs.Fields("金利グループ名").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ内容").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ番号").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ解約日").Value), _
            P8.FCStr(wRs.Fields("実行日").Value), _
            P8.FCStr(wRs.Fields("初回返済年月").Value), P8.FCStr(wRs.Fields("初回返済年月日").Value), _
            P8.FCStr(wRs.Fields("最終返済年月").Value), P8.FCStr(wRs.Fields("最終返済年月日").Value), _
            P8.FCStr(wRs.Fields("解約年月日").Value), _
            P8.FCDbl(wRs.Fields("初回返済額").Value), P8.FCDbl(wRs.Fields("毎月返済額").Value), P8.FCDbl(wRs.Fields("最終返済額").Value), _
            P8.FCStr(wRs.Fields("変動利率フラグ").Value), P8.FCDbl(wRs.Fields("利率").Value), _
            P8.FCDbl(wRs.Fields("返済単位月数").Value), _
            P8.FCDbl(wRs.Fields("融資金額").Value), _
            P8.FCDbl(wRs.Fields("融資残高").Value)

    wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
    
    DoEvents
    
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_借入一覧表_全件
'------------------------------------------------
Private Sub MX040_借入一覧表_全件(pCsvFileName As String)
'
    Dim ws01 As String
    Dim wsNendo As String, wsTbl As String
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wsTbl = "DBDA010_借入金"
    If GRpt.帳票名 <> "借入一覧表" And GRpt.帳票名 <> "借入一覧表(全件)" Then
        wsTbl = "DBDA010_貸付金"
    End If
'
    wstr = "Select "
    If GRpt.帳票名 = "借入一覧表" Then
        wstr = wstr & "K.借入番号,"
        wstr = wstr & "K.借入内容,"
    ElseIf GRpt.帳票名 = "借入一覧表(全件)" Then
        wstr = wstr & "K.借入番号,"
        wstr = wstr & "K.借入内容,"
    
    Else
        wstr = wstr & "K.借入番号 As 貸付番号,"
        wstr = wstr & "K.借入内容 As 貸付内容,"
    End If
    
    wstr = wstr & " KS.借入金種別名,"
    wstr = wstr & " K.銀行番号,"
    wstr = wstr & " G.銀行名,"
    wstr = wstr + "B.部門名,"
    
    wstr = wstr & "IIF(K.営業日区分 = " & P8.FCDbl(XMXA020_区分("営業日", "翌営業日")) & ",'翌営業日','前営業日') As 営業日,"
    wstr = wstr & "IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As 利息区分,"
    wstr = wstr & "IIF(K.利息計算日数区分 = " & P8.FCDbl(XMXA020_区分("利息計算日数", "営業日数")) & ",'営業日数','固定日数') As 利息計算日数,"
    wstr = wstr & "IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As 金利種別,"
    wstr = wstr & "KK.基準金利名,"
    wstr = wstr & "K.金利条件 As 金利備考,"
    
    wstr = wstr & "IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期借入金','長期借入金') AS 長短区分,"
    wstr = wstr & "IIF(K.有担保フラグ = " & P8.FCDbl(XMXA020_区分("有担フラグ", "無担保")) & ",'無担保','有担保') As 担保区分,"
    wstr = wstr & "K.担保名 As 担保内容,"
    wstr = wstr & "IIF(設備フラグ=1,'設備','運転資金') As 設備区分,"
    wstr = wstr & "K.資金用途,"
    
    wstr = wstr & " KG.金利グループ名,"
    wstr = wstr & " IIF(K.sm区分=1 and K.金融リストラ番号<>'','借入SM',IIF(K.sm区分=0 and not (K.金融解約実行日 is null),'解約SM','')) As ｼﾐｭﾚｰｼｮﾝ内容,"
    wstr = wstr & " K.金融リストラ番号 As ｼﾐｭﾚｰｼｮﾝ番号,"
    wstr = wstr & " format(K.金融解約実行日,'yyyy/mm/dd') As ｼﾐｭﾚｰｼｮﾝ解約日,"
    'wstr = wstr & "Format(K.保証料率,'#,##0.00000') As 保証料率,"
'    wstr = wstr & "Format(K.実行日,'yyyy/mm/dd') As 実行日,"
'    wstr = wstr & "Format(K.初回返済年月,'yyyy/mm/dd') As 初回返済年月,"
'    wstr = wstr & "Format(K.初回返済実行日,'yyyy/mm/dd') As 初回返済年月日,"
'    wstr = wstr & "Format(K.最終返済年月,'yyyy/mm/dd') As 最終返済年月,"
'    wstr = wstr & "Format(K.最終返済実行日,'yyyy/mm/dd') As 最終返済年月日,"
'    wstr = wstr & "Format(K.解約実行日,'yyyy/mm/dd') As 解約年月日,"
    
    '2012.10.24　追加 by k.kuni
    wstr = wstr & "Format(K.実行日,'" & Gfmtcsv年月日 & "') As 実行日,"
    wstr = wstr & "Format(K.初回返済年月,'" & Gfmtcsv年月日 & "') As 初回返済年月,"
    wstr = wstr & "Format(K.初回返済実行日,'" & Gfmtcsv年月日 & "') As 初回返済年月日,"
    wstr = wstr & "Format(K.最終返済年月,'" & Gfmtcsv年月日 & "') As 最終返済年月,"
    wstr = wstr & "Format(K.最終返済実行日,'" & Gfmtcsv年月日 & "') As 最終返済年月日,"
    wstr = wstr & "Format(K.解約実行日,'" & Gfmtcsv年月日 & "') As 解約年月日,"
    
    wstr = wstr & "IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.初回返済額,0) As 初回返済額,"
    wstr = wstr & "IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.毎月返済額,0) As 毎月返済額,"
    wstr = wstr & "IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.最終返済額,0) As 最終返済額,"
    
    wstr = wstr & "format(K.利率,'#,##0.00000') As 利率,"
    wstr = wstr & "K.返済単位月数,"
    
    wstr = wstr & "K.融資金額 As 融資金額"
    
    wstr = wstr & " From ((((" & wsTbl & " As K"
    wstr = wstr + " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr + " ON K.銀行番号 = G.銀行番号)"
    wstr = wstr + " Inner Join DAAA116_借入金種別 As KS"
    wstr = wstr + " ON K.借入金種別区分 = KS.借入金種別区分)"
    wstr = wstr + " LEFT JOIN DAAA200_部門マスタ AS B"
    wstr = wstr + "  ON B.部門番号 = K.プロジェクト番号)"
    wstr = wstr + " Left Join DAAA116_基準金利 As KK"
    wstr = wstr + " ON K.基準金利区分 = KK.基準金利区分)"
    wstr = wstr + " Left Join DAAA115_金利シミュレーショングループ As KG"
    wstr = wstr + " ON K.金利グループ区分 = KG.金利グループ区分"
    
    wstr = wstr + " Order BY K.銀行番号,K.借入番号"
    
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        '名称
        Write #1, _
            "借入番号", "借入内容", "借入金種別名", "銀行番号", "銀行名", "部門名", _
            "営業日", "利息区分", "利息計算日数", "金利種別", "基準金利名", "金利備考", _
            "長短区分", "担保区分", "担保内容", "設備区分", "資金用途", _
            "金利グループ名", "ｼﾐｭﾚｰｼｮﾝ内容", "ｼﾐｭﾚｰｼｮﾝ番号", "ｼﾐｭﾚｰｼｮﾝ解約日", _
            "実行日", "初回返済年月", "初回返済年月日", "最終返済年月", "最終返済年月日", "解約年月日", _
            "初回返済額", "毎月返済額", "最終返済額", _
            "利率", _
            "返済単位月数", _
            "融資金額"
            
    Do Until wRs.eof
        Write #1, _
            P8.FCStr(wRs.Fields("借入番号").Value), P8.FCStr(wRs.Fields("借入内容").Value), _
            P8.FCStr(wRs.Fields("借入金種別名").Value), _
            P8.FCStr(wRs.Fields("銀行番号").Value), P8.FCStr(wRs.Fields("銀行名").Value), P8.FCStr(wRs.Fields("部門名").Value), _
            P8.FCStr(wRs.Fields("営業日").Value), P8.FCStr(wRs.Fields("利息区分").Value), P8.FCStr(wRs.Fields("利息計算日数").Value), _
            P8.FCStr(wRs.Fields("金利種別").Value), P8.FCStr(wRs.Fields("基準金利名").Value), P8.FCStr(wRs.Fields("金利備考").Value), _
            P8.FCStr(wRs.Fields("長短区分").Value), P8.FCStr(wRs.Fields("担保区分").Value), P8.FCStr(wRs.Fields("担保内容").Value), _
            P8.FCStr(wRs.Fields("設備区分").Value), P8.FCStr(wRs.Fields("資金用途").Value), _
            P8.FCStr(wRs.Fields("金利グループ名").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ内容").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ番号").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ解約日").Value), _
            P8.FCStr(wRs.Fields("実行日").Value), _
            P8.FCStr(wRs.Fields("初回返済年月").Value), P8.FCStr(wRs.Fields("初回返済年月日").Value), _
            P8.FCStr(wRs.Fields("最終返済年月").Value), P8.FCStr(wRs.Fields("最終返済年月日").Value), _
            P8.FCStr(wRs.Fields("解約年月日").Value), _
            P8.FCDbl(wRs.Fields("初回返済額").Value), P8.FCDbl(wRs.Fields("毎月返済額").Value), P8.FCDbl(wRs.Fields("最終返済額").Value), _
            P8.FCDbl(wRs.Fields("利率").Value), _
            P8.FCDbl(wRs.Fields("返済単位月数").Value), _
            P8.FCDbl(wRs.Fields("融資金額").Value)

    wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
    
    DoEvents
    
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_平均金利平均残高推移表
'------------------------------------------------
Private Sub MX040_平均金利平均残高推移表(pCsvFileName As String)
'
    Dim j As Integer, wML As Integer
    Dim ws01 As String, ws02 As String, wsNendo As String
    Dim wsTbl As String
    
    Dim wsKin(12) As String
    Dim wsZan(12) As String, wsHzan(12) As String, wsRHeizan(12) As String
    Dim wsNisu(12) As String, wsRisoku(12) As String
    
    Dim wdKin(12) As Double
    Dim wdzan(12) As Double, wdHzan(12) As Double, wdRHeizan(12) As Double
    Dim wdNisu(12) As Double, wdRisoku(12) As Double
'
    On Error Resume Next
'
    'フィールド数
    Select Case GRpt.推移
    Case "月次", "四半期"
        wML = 12
    Case "年次", "半期"
        wML = 10
    End Select

    wsTbl = "DBDA010_借入金"
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wstr = "SELECT "
    wstr = wstr & "Z.借入番号,"
    wstr = wstr & " K.借入内容,"
    wstr = wstr & " KS.借入金種別名,"
    wstr = wstr & " K.銀行番号,"
    wstr = wstr & " G.銀行名,"
    wstr = wstr & "B.部門名,"
    
    wstr = wstr & "IIF(K.営業日区分 = " & P8.FCDbl(XMXA020_区分("営業日", "翌営業日")) & ",'翌営業日','前営業日') As 営業日,"
    wstr = wstr & "IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As 利息区分,"
    wstr = wstr & "IIF(K.利息計算日数区分 = " & P8.FCDbl(XMXA020_区分("利息計算日数", "営業日数")) & ",'営業日数','固定日数') As 利息計算日数,"
    wstr = wstr & "IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As 金利種別,"
    wstr = wstr & "KK.基準金利名,"
    wstr = wstr & "K.金利条件 As 金利備考,"
    
    wstr = wstr & "IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期借入金','長期借入金') AS 長短区分,"
    wstr = wstr & "IIF(K.有担保フラグ = " & P8.FCDbl(XMXA020_区分("有担フラグ", "無担保")) & ",'無担保','有担保') As 担保区分,"
    wstr = wstr & "K.担保名 As 担保内容,"
    wstr = wstr & "IIF(設備フラグ=1,'設備','運転資金') As 設備区分,"
    wstr = wstr & "K.資金用途,"
    
    wstr = wstr & "KG.金利グループ名,"
    wstr = wstr & "IIF(K.sm区分=1 and K.金融リストラ番号<>'','借入SM',IIF(K.sm区分=0 and not (K.金融解約実行日 is null),'解約SM','')) As ｼﾐｭﾚｰｼｮﾝ内容,"
    wstr = wstr & "K.金融リストラ番号 As ｼﾐｭﾚｰｼｮﾝ番号,"
    wstr = wstr & "format(K.金融解約実行日,'yyyy/mm/dd') As ｼﾐｭﾚｰｼｮﾝ解約日,"
    
    'wstr = wstr & "Format(K.実行日,'yyyy/mm/dd') As 実行日,"
    'wstr = wstr & "Format(K.初回返済年月,'yyyy/mm/dd') As 初回返済年月,"
    'wstr = wstr & "Format(K.初回返済実行日,'yyyy/mm/dd') As 初回返済年月日,"
    'wstr = wstr & "Format(K.最終返済年月,'yyyy/mm/dd') As 最終返済年月,"
    'wstr = wstr & "Format(K.最終返済実行日,'yyyy/mm/dd') As 最終返済年月日,"
    'wstr = wstr & "Format(K.解約実行日,'yyyy/mm/dd') As 解約年月日,"
    
    '2012.10.23　追加 by k.kunita
    wstr = wstr & "Format(K.実行日,'" & Gfmtcsv年月日 & "') As 実行日,"
    wstr = wstr & "Format(K.初回返済年月,'" & Gfmtcsv年月日 & "') As 初回返済年月,"
    wstr = wstr & "Format(K.初回返済実行日,'" & Gfmtcsv年月日 & "') As 初回返済年月日,"
    wstr = wstr & "Format(K.最終返済年月,'" & Gfmtcsv年月日 & "') As 最終返済年月,"
    wstr = wstr & "Format(K.最終返済実行日,'" & Gfmtcsv年月日 & "') As 最終返済年月日,"
    wstr = wstr & "Format(K.解約実行日,'" & Gfmtcsv年月日 & "') As 解約年月日,"
    
    wstr = wstr & "Format(IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.初回返済額,0),'#,##0') As 初回返済額,"
    wstr = wstr & "Format(IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.毎月返済額,0),'#,##0') As 毎月返済額,"
    wstr = wstr & "Format(IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.最終返済額,0),'#,##0') As 最終返済額,"
    
    'wstr = wstr + "IIF(K.変動利率フラグ=1,'*','') As 変動利率フラグ,"
    'wstr = wstr & "K.利率 As 利率,"
    wstr = wstr & "K.返済単位月数,"
    wstr = wstr & "K.融資金額 As 融資金額,"
    
    wstr = wstr & "残高合計,"
    wstr = wstr & "平均残高合計,"
    wstr = wstr & "利息計算平均残高合計,"
    wstr = wstr & "平均残高日数合計,"
    wstr = wstr & "平均利息基礎額合計,"
        
    For j = 1 To wML - 1
        ws01 = "_" & Right("0" & CStr(j), 2)

        wstr = wstr & "残高" & ws01 & ","
        wstr = wstr & "平均残高" & ws01 & ","
        wstr = wstr & "利息計算平均残高" & ws01 & ","
        wstr = wstr & "平均残高日数" & ws01 & ","
        wstr = wstr & "平均利息基礎額" & ws01 & ","
    Next j

    For j = wML To wML
        ws01 = "_" & Right("0" & CStr(j), 2)

        wstr = wstr & "残高" & ws01 & ","
        wstr = wstr & "平均残高" & ws01 & ","
        wstr = wstr & "利息計算平均残高" & ws01 & ","
        wstr = wstr & "平均残高日数" & ws01 & ","
        wstr = wstr & "平均利息基礎額" & ws01
    Next j

    wstr = wstr + " FROM ((((((DCDA010_借入残高推移表結果2 As Z"
    wstr = wstr + " Inner Join " & wsTbl & " As K"
    wstr = wstr + " ON Z.借入番号 = K.借入番号)"
    wstr = wstr & " INNER JOIN DCDA010_借入残高推移表結果 As Z1"
    wstr = wstr & " ON Z1.借入番号=K.借入番号)"
    wstr = wstr + " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr + " ON K.銀行番号 = G.銀行番号)"
    wstr = wstr + " LEFT JOIN DAAA200_部門マスタ AS B"
    wstr = wstr + " ON K.プロジェクト番号 = B.部門番号)"
    wstr = wstr + " Left Join DAAA116_借入金種別 As KS"
    wstr = wstr + " ON K.借入金種別区分 = KS.借入金種別区分)"
    wstr = wstr + " Left Join DAAA116_基準金利 As KK"
    wstr = wstr + " ON K.基準金利区分 = KK.基準金利区分)"
    wstr = wstr + " Left Join DAAA115_金利シミュレーショングループ As KG"
    wstr = wstr + " ON K.金利グループ区分 = KG.金利グループ区分"
    
    'Order
    wstr = wstr & " Order By K.銀行番号,Z.借入番号"
    
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        
        For j = 1 To wML
            ws01 = "_" & Right("0" & CStr(j), 2)
    
            If G基本情報.日付入力区分 = "0" Then
            '和暦入力
                wsZan(j) = w推移表タイトル.X番目年月(j) & "融資残高"
                wsKin(j) = w推移表タイトル.X番目年月(j) & "平均利率"
                wsHzan(j) = w推移表タイトル.X番目年月(j) & "平均残高"
                wsRHeizan(j) = w推移表タイトル.X番目年月(j) & "平残基礎"
                wsNisu(j) = w推移表タイトル.X番目年月(j) & "日数"
                wsRisoku(j) = w推移表タイトル.X番目年月(j) & "利息基礎"
                
            Else
            '西暦入力
                wsZan(j) = Format(w推移表タイトル.X番目年月(j), "yyyymm") & "融資残高"
                wsKin(j) = Format(w推移表タイトル.X番目年月(j), "yyyymm") & "平均利率"
                wsHzan(j) = Format(w推移表タイトル.X番目年月(j), "yyyymm") & "平均残高"
                wsRHeizan(j) = Format(w推移表タイトル.X番目年月(j), "yyyymm") & "平残基礎"
                wsNisu(j) = Format(w推移表タイトル.X番目年月(j), "yyyymm") & "日数"
                wsRisoku(j) = Format(w推移表タイトル.X番目年月(j), "yyyymm") & "利息基礎"
            End If
    
        Next j
        
        '名称
        If wML = 12 Then
            Write #1, _
                "借入番号", "借入内容", "借入金種別名", "銀行番号", "銀行名", "部門名", _
                "営業日", "利息区分", "利息計算日数", "金利種別", "基準金利名", "金利備考", _
                "長短区分", "担保区分", "担保内容", "設備区分", "資金用途", _
                "金利グループ名", "ｼﾐｭﾚｰｼｮﾝ内容", "ｼﾐｭﾚｰｼｮﾝ番号", "ｼﾐｭﾚｰｼｮﾝ解約日", _
                "実行日", "初回返済年月", "初回返済年月日", "最終返済年月", "最終返済年月日", "解約年月日", _
                "初回返済額", "毎月返済額", "最終返済額", "返済単位月数", "融資金額", _
                "合計融資残高", "合計平均金利", "合計平均残高", "合計平残基礎", "合計日数", "合計利息基礎", _
                wsZan(1), wsKin(1), wsHzan(1), wsRHeizan(1), wsNisu(1), wsRisoku(1), _
                wsZan(2), wsKin(2), wsHzan(2), wsRHeizan(2), wsNisu(2), wsRisoku(2), _
                wsZan(3), wsKin(3), wsHzan(3), wsRHeizan(3), wsNisu(3), wsRisoku(3), _
                wsZan(4), wsKin(4), wsHzan(4), wsRHeizan(4), wsNisu(4), wsRisoku(4), _
                wsZan(5), wsKin(5), wsHzan(5), wsRHeizan(5), wsNisu(5), wsRisoku(5), _
                wsZan(6), wsKin(6), wsHzan(6), wsRHeizan(6), wsNisu(6), wsRisoku(6), _
                wsZan(7), wsKin(7), wsHzan(7), wsRHeizan(7), wsNisu(7), wsRisoku(7), _
                wsZan(8), wsKin(8), wsHzan(8), wsRHeizan(8), wsNisu(8), wsRisoku(8), _
                wsZan(9), wsKin(9), wsHzan(9), wsRHeizan(9), wsNisu(9), wsRisoku(9), _
                wsZan(10), wsKin(10), wsHzan(10), wsRHeizan(10), wsNisu(10), wsRisoku(10), _
                wsZan(11), wsKin(11), wsHzan(11), wsRHeizan(11), wsNisu(11), wsRisoku(11), _
                wsZan(12), wsKin(12), wsHzan(12), wsRHeizan(12), wsNisu(12), wsRisoku(12)
        
        ElseIf wML = 10 Then
            Write #1, _
                "借入番号", "借入内容", "借入金種別名", "銀行番号", "銀行名", "部門名", _
                "営業日", "利息区分", "利息計算日数", "金利種別", "基準金利名", "金利備考", _
                "長短区分", "担保区分", "担保内容", "設備区分", "資金用途", _
                "金利グループ名", "ｼﾐｭﾚｰｼｮﾝ内容", "ｼﾐｭﾚｰｼｮﾝ番号", "ｼﾐｭﾚｰｼｮﾝ解約日", _
                "実行日", "初回返済年月", "初回返済年月日", "最終返済年月", "最終返済年月日", "解約年月日", _
                "初回返済額", "毎月返済額", "最終返済額", "返済単位月数", "融資金額", _
                "合計融資残高", "合計平均金利", "合計平均残高", "合計平残基礎", "合計日数", "合計利息基礎", _
                wsZan(1), wsKin(1), wsHzan(1), wsRHeizan(1), wsNisu(1), wsRisoku(1), _
                wsZan(2), wsKin(2), wsHzan(2), wsRHeizan(2), wsNisu(2), wsRisoku(2), _
                wsZan(3), wsKin(3), wsHzan(3), wsRHeizan(3), wsNisu(3), wsRisoku(3), _
                wsZan(4), wsKin(4), wsHzan(4), wsRHeizan(4), wsNisu(4), wsRisoku(4), _
                wsZan(5), wsKin(5), wsHzan(5), wsRHeizan(5), wsNisu(5), wsRisoku(5), _
                wsZan(6), wsKin(6), wsHzan(6), wsRHeizan(6), wsNisu(6), wsRisoku(6), _
                wsZan(7), wsKin(7), wsHzan(7), wsRHeizan(7), wsNisu(7), wsRisoku(7), _
                wsZan(8), wsKin(8), wsHzan(8), wsRHeizan(8), wsNisu(8), wsRisoku(8), _
                wsZan(9), wsKin(9), wsHzan(9), wsRHeizan(9), wsNisu(9), wsRisoku(9), _
                wsZan(10), wsKin(10), wsHzan(10), wsRHeizan(10), wsNisu(10), wsRisoku(10)
        End If
        
    
    Do Until wRs.eof
        
        wdzan(0) = P8.FCDbl(wRs("残高合計"))
        wdNisu(0) = P8.FCDbl(wRs("平均残高日数合計"))
        
        wdHzan(0) = P8.FCDbl(wRs("平均残高合計"))
        wdHzan(0) = P8.FRound(P8.FCDiv(wdHzan(0), wdNisu(0)), 0)
        
        wdRHeizan(0) = P8.FCDbl(wRs("利息計算平均残高合計"))
        wdRHeizan(0) = P8.FRound(P8.FCDiv(wdRHeizan(0), wdNisu(0)), 0)
        
        wdRisoku(0) = P8.FCDbl(wRs("平均利息基礎額合計"))

        '平均金利
        wdKin(0) = Format(Round(P8.FCDiv(wdRisoku(0) * 365 / wdNisu(0), wdRHeizan(0)) * 100, 6), "#,##0.00000")
        
        For j = 1 To wML
            ws01 = "_" & Right("0" & CStr(j), 2)
            
            wdzan(j) = P8.FCDbl(wRs("残高" & ws01))  '融資残高
            wdNisu(j) = P8.FCDbl(wRs("平均残高日数" & ws01))  '平均残高日数
            
            wdHzan(j) = P8.FCDbl(wRs("平均残高" & ws01))
            wdHzan(j) = P8.FRound(P8.FCDiv(wdHzan(j), wdNisu(j)), 0)
            
            wdRHeizan(j) = P8.FCDbl(wRs("利息計算平均残高" & ws01))
            wdRHeizan(j) = P8.FRound(P8.FCDiv(wdRHeizan(j), wdNisu(j)), 0)
            
            wdRisoku(j) = P8.FCDbl(wRs("平均利息基礎額" & ws01))
    
            '平均金利
            wdKin(j) = Format(Round(P8.FCDiv(wdRisoku(j) * 365 / wdNisu(j), wdRHeizan(j)) * 100, 6), "#,##0.00000")
        Next j
        
        If wML = 12 Then
            Write #1, _
                P8.FCStr(wRs.Fields("借入番号").Value), P8.FCStr(wRs.Fields("借入内容").Value), P8.FCStr(wRs.Fields("借入金種別名").Value), P8.FCStr(wRs.Fields("銀行番号").Value), P8.FCStr(wRs.Fields("銀行名").Value), P8.FCStr(wRs.Fields("部門名").Value), _
                P8.FCStr(wRs.Fields("営業日").Value), P8.FCStr(wRs.Fields("利息区分").Value), P8.FCStr(wRs.Fields("利息計算日数").Value), P8.FCStr(wRs.Fields("金利種別").Value), P8.FCStr(wRs.Fields("基準金利名").Value), P8.FCStr(wRs.Fields("金利備考").Value), _
                P8.FCStr(wRs.Fields("長短区分").Value), P8.FCStr(wRs.Fields("担保区分").Value), P8.FCStr(wRs.Fields("担保内容").Value), P8.FCStr(wRs.Fields("設備区分").Value), P8.FCStr(wRs.Fields("資金用途").Value), _
                P8.FCStr(wRs.Fields("金利グループ名").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ内容").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ番号").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ解約日").Value), _
                P8.FCStr(wRs.Fields("実行日").Value), P8.FCStr(wRs.Fields("初回返済年月").Value), P8.FCStr(wRs.Fields("初回返済年月日").Value), P8.FCStr(wRs.Fields("最終返済年月").Value), P8.FCStr(wRs.Fields("最終返済年月日").Value), P8.FCStr(wRs.Fields("解約年月日").Value), _
                P8.FCDbl(wRs.Fields("初回返済額").Value), P8.FCDbl(wRs.Fields("毎月返済額").Value), P8.FCDbl(wRs.Fields("最終返済額").Value), P8.FCDbl(wRs.Fields("返済単位月数").Value), P8.FCDbl(wRs.Fields("融資金額").Value), _
                wdzan(0), wdKin(0), wdHzan(0), wdRHeizan(0), wdNisu(0), wdRisoku(0), _
                wdzan(1), wdKin(1), wdHzan(1), wdRHeizan(1), wdNisu(1), wdRisoku(1), _
                wdzan(2), wdKin(2), wdHzan(2), wdRHeizan(2), wdNisu(2), wdRisoku(2), _
                wdzan(3), wdKin(3), wdHzan(3), wdRHeizan(3), wdNisu(3), wdRisoku(3), _
                wdzan(4), wdKin(4), wdHzan(4), wdRHeizan(4), wdNisu(4), wdRisoku(4), _
                wdzan(5), wdKin(5), wdHzan(5), wdRHeizan(5), wdNisu(5), wdRisoku(5), _
                wdzan(6), wdKin(6), wdHzan(6), wdRHeizan(6), wdNisu(6), wdRisoku(6), _
                wdzan(7), wdKin(7), wdHzan(7), wdRHeizan(7), wdNisu(7), wdRisoku(7), _
                wdzan(8), wdKin(8), wdHzan(8), wdRHeizan(8), wdNisu(8), wdRisoku(8), _
                wdzan(9), wdKin(9), wdHzan(9), wdRHeizan(9), wdNisu(9), wdRisoku(9), _
                wdzan(10), wdKin(10), wdHzan(10), wdRHeizan(10), wdNisu(10), wdRisoku(10), _
                wdzan(11), wdKin(11), wdHzan(11), wdRHeizan(11), wdNisu(11), wdRisoku(11), _
                wdzan(12), wdKin(12), wdHzan(12), wdRHeizan(12), wdNisu(12), wdRisoku(12)
        
        ElseIf wML = 10 Then
            Write #1, _
                P8.FCStr(wRs.Fields("借入番号").Value), P8.FCStr(wRs.Fields("借入内容").Value), P8.FCStr(wRs.Fields("借入金種別名").Value), P8.FCStr(wRs.Fields("銀行番号").Value), P8.FCStr(wRs.Fields("銀行名").Value), P8.FCStr(wRs.Fields("部門名").Value), _
                P8.FCStr(wRs.Fields("営業日").Value), P8.FCStr(wRs.Fields("利息区分").Value), P8.FCStr(wRs.Fields("利息計算日数").Value), P8.FCStr(wRs.Fields("金利種別").Value), P8.FCStr(wRs.Fields("基準金利名").Value), P8.FCStr(wRs.Fields("金利備考").Value), _
                P8.FCStr(wRs.Fields("長短区分").Value), P8.FCStr(wRs.Fields("担保区分").Value), P8.FCStr(wRs.Fields("担保内容").Value), P8.FCStr(wRs.Fields("設備区分").Value), P8.FCStr(wRs.Fields("資金用途").Value), _
                P8.FCStr(wRs.Fields("金利グループ名").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ内容").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ番号").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ解約日").Value), _
                P8.FCStr(wRs.Fields("実行日").Value), P8.FCStr(wRs.Fields("初回返済年月").Value), P8.FCStr(wRs.Fields("初回返済年月日").Value), P8.FCStr(wRs.Fields("最終返済年月").Value), P8.FCStr(wRs.Fields("最終返済年月日").Value), P8.FCStr(wRs.Fields("解約年月日").Value), _
                P8.FCDbl(wRs.Fields("初回返済額").Value), P8.FCDbl(wRs.Fields("毎月返済額").Value), P8.FCDbl(wRs.Fields("最終返済額").Value), P8.FCDbl(wRs.Fields("返済単位月数").Value), P8.FCDbl(wRs.Fields("融資金額").Value), _
                wdzan(0), wdKin(0), wdHzan(0), wdRHeizan(0), wdNisu(0), wdRisoku(0), _
                wdzan(1), wdKin(1), wdHzan(1), wdRHeizan(1), wdNisu(1), wdRisoku(1), _
                wdzan(2), wdKin(2), wdHzan(2), wdRHeizan(2), wdNisu(2), wdRisoku(2), _
                wdzan(3), wdKin(3), wdHzan(3), wdRHeizan(3), wdNisu(3), wdRisoku(3), _
                wdzan(4), wdKin(4), wdHzan(4), wdRHeizan(4), wdNisu(4), wdRisoku(4), _
                wdzan(5), wdKin(5), wdHzan(5), wdRHeizan(5), wdNisu(5), wdRisoku(5), _
                wdzan(6), wdKin(6), wdHzan(6), wdRHeizan(6), wdNisu(6), wdRisoku(6), _
                wdzan(7), wdKin(7), wdHzan(7), wdRHeizan(7), wdNisu(7), wdRisoku(7), _
                wdzan(8), wdKin(8), wdHzan(8), wdRHeizan(8), wdNisu(8), wdRisoku(8), _
                wdzan(9), wdKin(9), wdHzan(9), wdRHeizan(9), wdNisu(9), wdRisoku(9), _
                wdzan(10), wdKin(10), wdHzan(10), wdRHeizan(10), wdNisu(10), wdRisoku(10)
        End If

    wRs.MoveNext
    Loop
    
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
    
    DoEvents
    
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_平均金利平均残高表
'------------------------------------------------
Private Sub MX040_平均金利平均残高表(pCsvFileName As String)
'
    Dim j As Integer
    Dim wd01 As Double
    Dim ws01 As String, wsNendo As String
    Dim wsTbl As String
    
    Dim wdKin As Double
    Dim Zan As Double, Hzan As Double, RHeizan As Double
    Dim Nisu As Double, Risoku As Double
'
    On Error Resume Next
'
    wsTbl = "DBDA010_借入金"
    If GRpt.帳票名 = "借入残高表" Then
        wsTbl = "DBDA010_借入金"
    ElseIf GRpt.帳票名 = "貸付残高表" Then
        wsTbl = "DBDA010_貸付金"
    End If
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    'GInt1パラメータ
    ws01 = "_" & Right("00" & CStr(GInt1), 2)
    
    wstr = "SELECT "
    wstr = wstr & "Z.借入番号,"
    wstr = wstr & "K.借入内容,"
    wstr = wstr & "KS.借入金種別名,"
    
    wstr = wstr & "K.銀行番号,"
    wstr = wstr & "G.銀行名,"
    wstr = wstr + "B.部門名,"
    
    wstr = wstr & "IIF(K.営業日区分 = " & P8.FCDbl(XMXA020_区分("営業日", "翌営業日")) & ",'翌営業日','前営業日') As 営業日,"
    wstr = wstr & "IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As 利息区分,"
    wstr = wstr & "IIF(K.利息計算日数区分 = " & P8.FCDbl(XMXA020_区分("利息計算日数", "営業日数")) & ",'営業日数','固定日数') As 利息計算日数,"
    wstr = wstr & "IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As 金利種別,"
    wstr = wstr & "KK.基準金利名,"
    wstr = wstr & "K.金利条件 As 金利備考,"
    
    wstr = wstr & "IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期借入金','長期借入金') AS 長短区分,"
    wstr = wstr & "IIF(K.有担保フラグ = " & P8.FCDbl(XMXA020_区分("有担フラグ", "無担保")) & ",'無担保','有担保') As 担保区分,"
    wstr = wstr & "K.担保名 As 担保内容,"
    wstr = wstr & "IIF(設備フラグ=1,'設備','運転資金') As 設備区分,"
    wstr = wstr & "K.資金用途,"
    
    wstr = wstr & "KG.金利グループ名,"
    wstr = wstr & "IIF(K.sm区分=1 and K.金融リストラ番号<>'','借入SM',IIF(K.sm区分=0 and not (K.金融解約実行日 is null),'解約SM','')) As ｼﾐｭﾚｰｼｮﾝ内容,"
    wstr = wstr & "K.金融リストラ番号 As ｼﾐｭﾚｰｼｮﾝ番号,"
'    wstr = wstr & "format(K.金融解約実行日,'yyyy/mm/dd') As ｼﾐｭﾚｰｼｮﾝ解約日,"
    
    'wstr = wstr & "Format(K.実行日,'yyyy/mm/dd') As 実行日,"
    'wstr = wstr & "Format(K.初回返済年月,'yyyy/mm/dd') As 初回返済年月,"
    'wstr = wstr & "Format(K.初回返済実行日,'yyyy/mm/dd') As 初回返済年月日,"
    'wstr = wstr & "Format(K.最終返済年月,'yyyy/mm/dd') As 最終返済年月,"
    'wstr = wstr & "Format(K.最終返済実行日,'yyyy/mm/dd') As 最終返済年月日,"
    'wstr = wstr & "Format(K.解約実行日,'yyyy/mm/dd') As 解約年月日,"
    
    '2012.10.23　追加 by k.kunita
    wstr = wstr & "format(K.金融解約実行日,'" & Gfmtcsv年月日 & "') As ｼﾐｭﾚｰｼｮﾝ解約日,"
    
    wstr = wstr & "Format(K.実行日,'" & Gfmtcsv年月日 & "') As 実行日,"
    wstr = wstr & "Format(K.初回返済年月,'" & Gfmtcsv年月日 & "') As 初回返済年月,"
    wstr = wstr & "Format(K.初回返済実行日,'" & Gfmtcsv年月日 & "') As 初回返済年月日,"
    wstr = wstr & "Format(K.最終返済年月,'" & Gfmtcsv年月日 & "') As 最終返済年月,"
    wstr = wstr & "Format(K.最終返済実行日,'" & Gfmtcsv年月日 & "') As 最終返済年月日,"
    wstr = wstr & "Format(K.解約実行日,'" & Gfmtcsv年月日 & "') As 解約年月日,"
    
    wstr = wstr & "IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.初回返済額,0) As 初回返済額,"
    wstr = wstr & "IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.毎月返済額,0) As 毎月返済額,"
    wstr = wstr & "IIF(K.手入力区分=" & P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) & ",K.最終返済額,0) As 最終返済額,"
    
    'wstr = wstr + "IIF(K.変動利率フラグ=1,'*','') As 変動利率フラグ,"
    'wstr = wstr & "K.利率 As 利率,"
    wstr = wstr & "K.返済単位月数,"
    
    wstr = wstr & "K.融資金額 As 融資金額,"
    
    wstr = wstr & "残高" & ws01 & ","
    wstr = wstr & "平均残高" & ws01 & ","
    wstr = wstr & "利息計算平均残高" & ws01 & ","
    wstr = wstr & "平均残高日数" & ws01 & ","
    wstr = wstr & "平均利息基礎額" & ws01
    
    wstr = wstr + " FROM ((((((DCDA010_借入残高推移表結果2 As Z"
    wstr = wstr + " Inner Join " & wsTbl & " As K"
    wstr = wstr + " ON Z.借入番号 = K.借入番号)"
    wstr = wstr & " INNER JOIN DCDA010_借入残高推移表結果 As Z1"
    wstr = wstr & " ON Z1.借入番号=K.借入番号)"
    wstr = wstr + " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr + " ON K.銀行番号 = G.銀行番号)"
    wstr = wstr + " LEFT JOIN DAAA200_部門マスタ AS B"
    wstr = wstr + " ON K.プロジェクト番号 = B.部門番号)"
    wstr = wstr + " Left Join DAAA116_借入金種別 As KS"
    wstr = wstr + " ON K.借入金種別区分 = KS.借入金種別区分)"
    wstr = wstr + " Left Join DAAA116_基準金利 As KK"
    wstr = wstr + " ON K.基準金利区分 = KK.基準金利区分)"
    wstr = wstr + " Left Join DAAA115_金利シミュレーショングループ As KG"
    wstr = wstr + " ON K.金利グループ区分 = KG.金利グループ区分"
    
    wstr = wstr & " Where 平均残高" + ws01 + "<>0"
    wstr = wstr & " Or 利息計算平均残高" + ws01 + "<>0"
    wstr = wstr & " Or 平均利息基礎額" + ws01 + "<>0"
    'wstr = wstr & " Where 残高" + ws01 + "<>0"
    
    wstr = wstr & " Order By K.銀行番号,Z.借入番号"
    
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
        
    On Error GoTo Err_Hundle
        '名称
        Write #1, _
            "借入番号", "借入内容", "借入金種別名", "銀行番号", "銀行名", "部門名", _
            "営業日", "利息区分", "利息計算日数", "金利種別", "基準金利名", "金利備考", _
            "長短区分", "担保区分", "担保内容", "設備区分", "資金用途", _
            "金利グループ名", "ｼﾐｭﾚｰｼｮﾝ内容", "ｼﾐｭﾚｰｼｮﾝ番号", "ｼﾐｭﾚｰｼｮﾝ解約日", _
            "実行日", "初回返済年月", "初回返済年月日", "最終返済年月", "最終返済年月日", "解約年月日", _
            "初回返済額", "毎月返済額", "最終返済額", "返済単位月数", "融資金額", _
            "融資残高", "平均金利", "平均残高", "平残基礎", "日数", "利息基礎"
    
    Do Until wRs.eof
        
        'GInt1パラメータ
        ws01 = "_" & Right("00" & CStr(GInt1), 2)
            
            Zan = P8.FCDbl(wRs("残高" & ws01))  '融資残高
            Nisu = P8.FCDbl(wRs("平均残高日数" & ws01))  '平均残高日数
            
            Hzan = P8.FCDbl(wRs("平均残高" & ws01))
            Hzan = P8.FRound(P8.FCDiv(Hzan, Nisu), 0)
            
            RHeizan = P8.FCDbl(wRs("利息計算平均残高" & ws01))
            RHeizan = P8.FRound(P8.FCDiv(RHeizan, Nisu), 0)
            
            Risoku = P8.FCDbl(wRs("平均利息基礎額" & ws01))
    
            '平均金利
            wdKin = Format(Round(P8.FCDiv(Risoku * 365 / Nisu, RHeizan) * 100, 6), "#,##0.00000")
        
        Write #1, _
            P8.FCStr(wRs.Fields("借入番号").Value), P8.FCStr(wRs.Fields("借入内容").Value), P8.FCStr(wRs.Fields("借入金種別名").Value), P8.FCStr(wRs.Fields("銀行番号").Value), P8.FCStr(wRs.Fields("銀行名").Value), P8.FCStr(wRs.Fields("部門名").Value), _
            P8.FCStr(wRs.Fields("営業日").Value), P8.FCStr(wRs.Fields("利息区分").Value), P8.FCStr(wRs.Fields("利息計算日数").Value), P8.FCStr(wRs.Fields("金利種別").Value), P8.FCStr(wRs.Fields("基準金利名").Value), P8.FCStr(wRs.Fields("金利備考").Value), _
            P8.FCStr(wRs.Fields("長短区分").Value), P8.FCStr(wRs.Fields("担保区分").Value), P8.FCStr(wRs.Fields("担保内容").Value), P8.FCStr(wRs.Fields("設備区分").Value), P8.FCStr(wRs.Fields("資金用途").Value), _
            P8.FCStr(wRs.Fields("金利グループ名").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ内容").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ番号").Value), P8.FCStr(wRs.Fields("ｼﾐｭﾚｰｼｮﾝ解約日").Value), _
            P8.FCStr(wRs.Fields("実行日").Value), P8.FCStr(wRs.Fields("初回返済年月").Value), P8.FCStr(wRs.Fields("初回返済年月日").Value), P8.FCStr(wRs.Fields("最終返済年月").Value), P8.FCStr(wRs.Fields("最終返済年月日").Value), P8.FCStr(wRs.Fields("解約年月日").Value), _
            P8.FCDbl(wRs.Fields("初回返済額").Value), P8.FCDbl(wRs.Fields("毎月返済額").Value), P8.FCDbl(wRs.Fields("最終返済額").Value), P8.FCDbl(wRs.Fields("返済単位月数").Value), P8.FCDbl(wRs.Fields("融資金額").Value), _
            Zan, wdKin, Hzan, RHeizan, Nisu, Risoku
            
    wRs.MoveNext
    Loop
    
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_固定資産台帳
'------------------------------------------------
Private Sub MX040_固定資産台帳(pCsvFileName As String)
'
    Dim ws01 As String
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wstr = ""
    wstr = wstr & "Select "
    
    wstr = wstr & "SK.設備番号,"
    wstr = wstr & "設備名,"
    wstr = wstr & "B.部門名,"
    wstr = wstr & " IIF(S.減価償却費区分 = '" & XMXA020_区分("減価償却費区分", "一般管理費") & "','一般管理費',"
    wstr = wstr & " IIF(S.減価償却費区分 = '" & XMXA020_区分("減価償却費区分", "製造原価") & "', '製造原価','')) As 減価償却費区分,"
    wstr = wstr & " IIF(S.資産区分 = '" & XMXA020_区分("資産区分", "有形資産") & "','有形資産',"
    wstr = wstr & " IIF(S.資産区分 = '" & XMXA020_区分("資産区分", "有形資産2") & "','有形資産2',"
    wstr = wstr & " IIF(S.資産区分 = '" & XMXA020_区分("資産区分", "無形資産") & "','無形資産',"
    wstr = wstr & " IIF(S.資産区分 = '" & XMXA020_区分("資産区分", "無形資産2") & "','無形資産2',"
    wstr = wstr & " IIF(S.資産区分 = '" & XMXA020_区分("資産区分", "損金設備") & "','損金設備',"
    wstr = wstr & " IIF(S.資産区分 = '" & XMXA020_区分("資産区分", "建物") & "','建物',"
    wstr = wstr & " IIF(S.資産区分 = '" & XMXA020_区分("資産区分", "土地") & "','土地',"
    wstr = wstr & " IIF(S.資産区分 = '" & XMXA020_区分("資産区分", "有価証券") & "','有価証券',"
    wstr = wstr & " IIF(S.資産区分 = '" & XMXA020_区分("資産区分", "その他投資") & "','その他投資',"
    wstr = wstr & " IIF(S.資産区分 = '" & XMXA020_区分("資産区分", "その他固定資産") & "','その他固定資産',"
    wstr = wstr & " IIF(S.資産区分 = '" & XMXA020_区分("資産区分", "繰延資産") & "','繰延資産',''))))))))))) AS 資産区分,"
    wstr = wstr & "K.勘定科目名,"
    wstr = wstr & "S.設備金額 As 設備金額,"
    wstr = wstr & "S.償却年数,"
    wstr = wstr & "Format(S.設備年月,'yyyy/mm/dd') As 設備年月,"
    wstr = wstr & "Format(S.設備購入年月日,'yyyy/mm/dd') As 設備購入年月日,"
    wstr = wstr & " IIf(S.償却区分 = '" + XMXA020_区分("償却区分", "定額法") + "','定額法',"
    wstr = wstr & " IIf(S.償却区分 = '" + XMXA020_区分("償却区分", "定率法") + "','定率法',"
    wstr = wstr & " IIf(S.償却区分 = '" + XMXA020_区分("償却区分", "均等償却") + "','均等償却',''))) AS 償却区分,"
    wstr = wstr & " Format(IIF(S.償却区分 = '" & XMXA020_区分("償却区分", "定率法") & "',SM.定率法償却率,"
    wstr = wstr & " SM.定額法償却率), '#,##0.000') As 旧償却率,"
    wstr = wstr & " Format(IIF(S.償却区分 = '" & XMXA020_区分("償却区分", "定率法") & "',SM.新定率法償却率,"
    wstr = wstr & " SM.新定額法償却率), '#,##0.000') As 新償却率,"
    
    wstr = wstr & "Format(SM.改定償却率, '#,##0.000') As 改定償却率,"
    wstr = wstr & "Format(SM.保証率, '#,##0.000') As 保証率,"
    wstr = wstr & "Format(S.残存率, '#,##0.000') As 残存率,"
    wstr = wstr & "S.支払サイト,"
    wstr = wstr & " IIF(S.課税区分 = '" & XMXA020_区分("課税区分", "不課税") & "','不課税',"
    wstr = wstr & " IIF(S.課税区分 = '" & XMXA020_区分("課税区分", "課税") & "','課税',"
    wstr = wstr & " IIF(S.課税区分 = '" & XMXA020_区分("課税区分", "非課税") & "', '非課税',''))) As 課税区分,"
    wstr = wstr & "Format(S.資産売却年月日,'yyyy/mm/dd') As 資産売却年月日,"
    wstr = wstr & "S.資産売却額 As 資産売却額,"
    wstr = wstr & "S.設備リストラ番号,"
    wstr = wstr & "S.回収サイト,"
    wstr = wstr & " IIF(S.売上課税区分 = '" & XMXA020_区分("課税区分", "不課税") & "','不課税',"
    wstr = wstr & " IIF(S.売上課税区分 = '" & XMXA020_区分("課税区分", "課税") & "','課税',"
    wstr = wstr & " IIF(S.売上課税区分 = '" & XMXA020_区分("課税区分", "非課税") & "','非課税',''))) As 売上課税区分,"
    
    wstr = wstr & "新規_01+調整償却_01 As 取得額,"
    wstr = wstr & "期首_01 As 期首簿価,"
    wstr = wstr & "償却_01+調整償却_01 As 当期償却額,"
    wstr = wstr & "特別償却_01 As 特別償却額,"
    wstr = wstr & "売却額_01 As 売却額,"
    wstr = wstr & "売却益_01-売却損_01 As 売却損益,"
    wstr = wstr & "残存_01 As 残存額"
    
    wstr = wstr & " INTO [Text;database=" & wCsvDir & "]." & "[" & pCsvFileName & "]"
    
    wstr = wstr & " FROM (((DCCA010_設備推移結果 As SK"
    wstr = wstr & " INNER JOIN DBCA010_設備計画 As S"
    wstr = wstr & " ON SK.設備番号 = S.設備番号)"
    wstr = wstr & " INNER JOIN DAAB010_償却率マスタ As SM"
    wstr = wstr & " ON S.償却年数 = SM.償却年数)"
    wstr = wstr & " LEFT JOIN DAAC020_固定資産部門マスタ As B"
    wstr = wstr & " ON S.部門番号 = B.部門番号)"
    wstr = wstr & " LEFT JOIN DAAC010_固定資産勘定科目マスタ As K"
    wstr = wstr & " ON S.勘定科目番号 = K.勘定科目番号"
    wstr = wstr & " Order By S.部門番号,S.資産区分,S.勘定科目番号,S.償却区分,SK.設備番号,S.設備年月"
    
    GDb.Execute wstr
    
    DoEvents
    
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
End Sub

'------------------------------------------------
' MX040_資金繰表
'------------------------------------------------
Public Sub MX040_資金繰表()
'
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim strFileName As String                             'ファイル名(フルパス)
    Dim strFilename2 As String
    Dim strSheetName As String                            'シート名

    Dim j As Integer
    Dim wiCnt As Integer, w分母 As Integer
'
    On Error Resume Next
'
    '---------------------
    ' EXCELファイルを開く
    '---------------------
    strFileName = GCurDir & "\資金繰表.xls"                     'ファイル名をセット
    strFilename2 = wCsvDir & "\" & Format(Date, "yyyymmdd") & "資金繰表.xls"

    If Dir(strFilename2) <> "" Then
        Kill strFilename2
    End If

    '----------------------
    ' EXCELファイルをコピー
    '----------------------
    FileCopy strFileName, strFilename2
    strSheetName = "資金繰表"                               'シート名をセット

    Set xlApp = CreateObject("Excel.Application")         'Application生成
    xlApp.Workbooks.Open FileName:=strFilename2, UpdateLinks:=0 'EXCELを開く
    xlApp.Visible = True                                       'EXCELの表示

    Set xlBook = xlApp.Workbooks(Dir(strFilename2))        'Workbook
    Set xlSheet = xlBook.Worksheets(strSheetName)         'Worksheet

    '--------------
    ' シートの編集
    '--------------
    xlSheet.Cells(2, 1).Value = "平成" & GRpt.テキスト_01 & "年度　資金繰表"

    If GRpt.千円単位 = 1 Then
        w分母 = 1000
        xlSheet.Cells(5, 1).Value = "（千円単位）"
    Else
        w分母 = 1
        xlSheet.Cells(5, 1).Value = "（円単位）"
    End If

    wiCnt = 11
    For j = 1 To 12
        xlSheet.Cells(5, wiCnt).Value = CStr(Format(w推移表タイトル.X番目年月(j), "mm月"))

        wiCnt = wiCnt + 12
    Next j

    wiCnt = 15
    For j = 1 To 12
        xlSheet.Cells(20, wiCnt - 4).Value = wdRisoku(j) + wdRisoku2(j) / w分母
        xlSheet.Cells(32, wiCnt - 4).Value = wdYushi(j) / w分母
        xlSheet.Cells(33, wiCnt - 4).Value = wdYushi2(j) / w分母
        xlSheet.Cells(35, wiCnt - 4).Value = wdGankin(j) / w分母
        xlSheet.Cells(36, wiCnt - 4).Value = wdGankin2(j) / w分母
        xlSheet.Cells(38, wiCnt - 4).Value = wdYZan(j) / w分母
        xlSheet.Cells(39, wiCnt - 4).Value = wdYZan2(j) / w分母
        
        xlSheet.Cells(20, wiCnt).Value = wdRisoku(j) + wdRisoku2(j) / w分母
        xlSheet.Cells(32, wiCnt).Value = wdYushi(j) / w分母
        xlSheet.Cells(33, wiCnt).Value = wdYushi2(j) / w分母
        xlSheet.Cells(35, wiCnt).Value = wdGankin(j) / w分母
        xlSheet.Cells(36, wiCnt).Value = wdGankin2(j) / w分母
        xlSheet.Cells(38, wiCnt).Value = wdYZan(j) / w分母
        xlSheet.Cells(39, wiCnt).Value = wdYZan2(j) / w分母

        wiCnt = wiCnt + 12
    Next j

    '-----------------------
    ' EXCELファイル終了処理
    '-----------------------
    xlBook.Close saveChanges:=True                       'ブックを保存して終了
    xlApp.Quit                                           'EXCELを閉じる

    Set xlSheet = Nothing                                'オブジェクトの解放
    Set xlBook = Nothing                                 'オブジェクトの解放
    Set xlApp = Nothing                                  'オブジェクトの解放
'
    If Err.Number Then
        GSstrt帳票Msg = "出力できませんでした"
    End If
'
    On Error GoTo 0
'
End Sub

'------------------------------------------------
' MX040_事業計画表
'------------------------------------------------
Public Sub MX040_事業計画表()
'
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim strFileName As String                             'ファイル名(フルパス)
    Dim strFilename2 As String
    Dim strSheetName As String                            'シート名
'
    On Error Resume Next
'
    '---------------------
    ' EXCELファイルを開く
    '---------------------
    strFileName = GCurDir & "\事業計画表.xls"                     'ファイル名をセット
    strFilename2 = wCsvDir & "\" & Format(Date, "yyyymmdd") & "事業計画表.xls"

    If Dir(strFilename2) <> "" Then
        Kill strFilename2
    End If

    '----------------------
    ' EXCELファイルをコピー
    '----------------------
    FileCopy strFileName, strFilename2
    strSheetName = "事業計画表"                               'シート名をセット

    Set xlApp = CreateObject("Excel.Application")         'Application生成
    xlApp.Workbooks.Open FileName:=strFilename2, UpdateLinks:=0 'EXCELを開く
    xlApp.Visible = True                                       'EXCELの表示

    Set xlBook = xlApp.Workbooks(Dir(strFilename2))        'Workbook
    Set xlSheet = xlBook.Worksheets(strSheetName)         'Worksheet

    '--------------
    ' シートの編集
    '--------------
    xlSheet.Cells(5, 5).Value = CStr("(" & w推移表タイトル.X番目年月(12) & "期)")
    xlSheet.Cells(23, 5).Value = wdYushi(0) + wdYushi2(0) / 1000

    '-----------------------
    ' EXCELファイル終了処理
    '-----------------------
    xlBook.Close saveChanges:=True                       'ブックを保存して終了
    xlApp.Quit                                           'EXCELを閉じる

    Set xlSheet = Nothing                                'オブジェクトの解放
    Set xlBook = Nothing                                 'オブジェクトの解放
    Set xlApp = Nothing                                  'オブジェクトの解放
'
    On Error GoTo 0
'
End Sub

'------------------------------------------------
' MX040_現状借換後対比表
'------------------------------------------------
Public Sub MX040_現状借換後対比表()
'
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim strFileName As String                             'ファイル名(フルパス)
    Dim strFilename2 As String
    Dim strSheetName As String                            'シート名

    Dim j As Integer, wML As Integer
    Dim wiCnt As Integer, w分母 As Integer
    Dim ws01 As String
'
    On Error Resume Next
'
    '---------------------
    ' EXCELファイルを開く
    '---------------------
    strFileName = GCurDir & "\現状借換後対比表.xls"                     'ファイル名をセット
    strFilename2 = wCsvDir & "\" & Format(Date, "yyyymmdd") & "現状借換後対比表.xls"

    If Dir(strFilename2) <> "" Then
        Kill strFilename2
    End If

    '----------------------
    ' EXCELファイルをコピー
    '----------------------
    FileCopy strFileName, strFilename2

    'シート名をセット
    If GRpt.推移 = "年次" Or GRpt.推移 = "半期" Then
        wML = 10
        strSheetName = "現状借換後対比表_2"
    ElseIf GRpt.推移 = "四半期" Or GRpt.推移 = "月次" Then
        wML = 12
        strSheetName = "現状借換後対比表"
    End If

    Set xlApp = CreateObject("Excel.Application")         'Application生成
    xlApp.Workbooks.Open FileName:=strFilename2, UpdateLinks:=0 'EXCELを開く
    xlApp.Visible = True                                       'EXCELの表示

    Set xlBook = xlApp.Workbooks(Dir(strFilename2))        'Workbook
    Set xlSheet = xlBook.Worksheets(strSheetName)         'Worksheet

    '--------------
    ' シートの編集
    '--------------
    ws01 = "借換計画番号：" & GRpt.コンボ_01
    ws01 = ws01 & ", 借換指定年月：" & GRpt.コンボ_02
    ws01 = ws01 & ", 推移表開始年度：" & GRpt.テキスト_01
    xlSheet.Cells(3, 1).Value = ws01

    If GRpt.千円単位 = 1 Then
        w分母 = 1000
        xlSheet.Cells(3, 14).Value = "（千円単位）"
    Else
        w分母 = 1
        xlSheet.Cells(3, 14).Value = "（円単位）"
    End If

    wiCnt = 4
    For j = 1 To wML
        xlSheet.Cells(4, wiCnt).Value = CStr(w推移表タイトル.X番目年月(j))

        wiCnt = wiCnt + 1
    Next j

    wiCnt = 3
    For j = 0 To wML
        xlSheet.Cells(5, wiCnt).Value = wdYushi(j) / w分母
        xlSheet.Cells(6, wiCnt).Value = wdGankin(j) / w分母
        xlSheet.Cells(7, wiCnt).Value = wdRisoku(j) / w分母
        xlSheet.Cells(8, wiCnt).Value = wdHensai(j) / w分母
        xlSheet.Cells(9, wiCnt).Value = wdKaiyaku(j) / w分母
        xlSheet.Cells(10, wiCnt).Value = wdYZan(j) / w分母

        xlSheet.Cells(11, wiCnt).Value = wdYushi2(j) / w分母
        xlSheet.Cells(12, wiCnt).Value = wdGankin2(j) / w分母
        xlSheet.Cells(13, wiCnt).Value = wdRisoku2(j) / w分母
        xlSheet.Cells(14, wiCnt).Value = wdHensai2(j) / w分母
        xlSheet.Cells(15, wiCnt).Value = wdKaiyaku2(j) / w分母
        xlSheet.Cells(16, wiCnt).Value = wdYZan2(j) / w分母

        wiCnt = wiCnt + 1
    Next j

    '-----------------------
    ' EXCELファイル終了処理
    '-----------------------
    xlBook.Close saveChanges:=True                       'ブックを保存して終了
    xlApp.Quit                                           'EXCELを閉じる

    Set xlSheet = Nothing                                'オブジェクトの解放
    Set xlBook = Nothing                                 'オブジェクトの解放
    Set xlApp = Nothing                                  'オブジェクトの解放
'
    On Error GoTo 0
'
End Sub
    
'------------------------------------------------
' MX040_借入残高推移表データ取得
'------------------------------------------------
Public Sub MX040_借入残高推移表データ取得(p推移表タイトル As MAA910_推移表タイトル, pSikin As Integer, pJigyo As Integer)
'
    Dim j As Integer, w番号 As String
    Dim ws01 As String
'
    On Error GoTo MX040_借入残高推移表データ取得_ERR
'
    w推移表タイトル = p推移表タイトル

    Nendo = GRpt.テキスト_01
    kubun = GRpt.推移
    UriNo = "": KarNo = ""
    LeaNo = "": SetNo = ""
    KinNo = "": SKinNo = ""
'
'----------< Create \金剛石CSV or \借換たろうCSV >----------------------------------
    ws01 = wCsvDir
    If Dir(ws01, vbDirectory) = "" Then
        MkDir (ws01)
        
        If Err.Number Then
            Err.Clear
            Exit Sub
        End If
    End If
'
    DoEvents
'
    For j = 0 To 12
        wdYushi(j) = 0
        wdGankin(j) = 0
        wdRisoku(j) = 0
        wdHensai(j) = 0
        wdKaiyaku(j) = 0
        wdYZan(j) = 0
        wdYushi2(j) = 0
        wdGankin2(j) = 0
        wdRisoku2(j) = 0
        wdHensai2(j) = 0
        wdKaiyaku2(j) = 0
        wdYZan2(j) = 0
    Next
'
    wstr = "Select K.長短区分,"
    For j = 1 To 12
        w番号 = Right("00" + CStr(j), 2)
        
        wstr = wstr & "sum(融資_" & w番号 & ") As 融資_" & w番号 & "集計,"
        wstr = wstr & "sum(元金_" & w番号 & ") As 元金_" & w番号 & "集計,"
        wstr = wstr & "sum(利息_" & w番号 & ") As 利息_" & w番号 & "集計,"
        wstr = wstr & "sum(返済_" & w番号 & ") As 返済_" & w番号 & "集計,"
        'wstr = wstr & "sum(解約_" & w番号 & ") As 解約_" & w番号 & "集計,"
        wstr = wstr & "sum(残高_" & w番号 & ") As 残高_" & w番号 & "集計,"
    Next

    wstr = wstr & "sum(融資合計) As 融資合計集計,"
    wstr = wstr & "sum(元金合計) As 元金合計集計,"
    wstr = wstr & "sum(利息合計) As 利息合計集計,"
    wstr = wstr & "sum(返済合計) As 返済合計集計,"
    'wstr = wstr & "sum(解約合計) As 解約合計集計,"
    wstr = wstr & "sum(残高合計) As 残高合計集計"
    wstr = wstr & " FROM DCDA010_借入残高推移表結果 As Z"
    wstr = wstr & " INNER JOIN DBDA010_借入金 As K"
    wstr = wstr & " ON Z.借入番号 = K.借入番号"
    wstr = wstr & " GROUP BY K.長短区分"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    Do Until wRs.eof
        If P8.FCDbl(wRs("長短区分")) = P8.FCDbl(XMXA020_区分("長短区分", "長期借入金")) Then
            For j = 1 To 12
                w番号 = Right("00" + CStr(j), 2)
                
                wdYushi2(j) = P8.FCDbl(wRs("融資_" & w番号 & "集計"))
                wdGankin2(j) = P8.FCDbl(wRs("元金_" & w番号 & "集計"))
                wdRisoku2(j) = P8.FCDbl(wRs("利息_" & w番号 & "集計"))
                wdHensai2(j) = P8.FCDbl(wRs("返済_" & w番号 & "集計"))
                wdYZan2(j) = P8.FCDbl(wRs("残高_" & w番号 & "集計"))
            Next
        
            wdYushi2(0) = P8.FCDbl(wRs("融資合計集計"))
            wdGankin2(0) = P8.FCDbl(wRs("元金合計集計"))
            wdRisoku2(0) = P8.FCDbl(wRs("利息合計集計"))
            wdHensai2(0) = P8.FCDbl(wRs("返済合計集計"))
            wdYZan2(0) = P8.FCDbl(wRs("残高合計集計"))
        Else
            For j = 1 To 12
                w番号 = Right("00" + CStr(j), 2)
                
                wdYushi(j) = P8.FCDbl(wRs("融資_" & w番号 & "集計"))
                wdGankin(j) = P8.FCDbl(wRs("元金_" & w番号 & "集計"))
                wdRisoku(j) = P8.FCDbl(wRs("利息_" & w番号 & "集計"))
                wdHensai(j) = P8.FCDbl(wRs("返済_" & w番号 & "集計"))
                wdYZan(j) = P8.FCDbl(wRs("残高_" & w番号 & "集計"))
            Next
        
            wdYushi(0) = P8.FCDbl(wRs("融資合計集計"))
            wdGankin(0) = P8.FCDbl(wRs("元金合計集計"))
            wdRisoku(0) = P8.FCDbl(wRs("利息合計集計"))
            wdHensai(0) = P8.FCDbl(wRs("返済合計集計"))
            wdYZan(0) = P8.FCDbl(wRs("残高合計集計"))
        End If
    
        wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
'
    If pSikin = True Then
        Call MX040_資金繰表
    End If
    If pJigyo = True Then
        Call MX040_事業計画表
    End If
'
    Exit Sub
'
'----------< ERROR ROUTINE >--------------------------------------------------------
MX040_借入残高推移表データ取得_ERR:
    pERR_MES = pPROGRAM_ID + "/ MX040_借入残高推移表データ取得() でエラー" + vbCrLf + vbCrLf + _
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
' MX040_借入残高推移表比較データ取得
'------------------------------------------------
Public Sub MX040_借入残高推移表比較データ取得(p推移表タイトル As MAA910_推移表タイトル, pHikaku As Integer)
'
    Dim j As Integer, w番号 As String
    Dim ws01 As String
'
    On Error GoTo MX040_借入残高推移表比較データ取得_ERR
'
    w推移表タイトル = p推移表タイトル

    Nendo = GRpt.テキスト_01
    kubun = GRpt.推移
    UriNo = "": KarNo = ""
    LeaNo = "": SetNo = ""
    KinNo = "": SKinNo = ""
'
'----------< Create \金剛石CSV or \借換たろうCSV >----------------------------------
    ws01 = wCsvDir
    If Dir(ws01, vbDirectory) = "" Then
        MkDir (ws01)
        
        If Err.Number Then
            Err.Clear
            Exit Sub
        End If
    End If
'
    DoEvents
'
    For j = 0 To 12
        wdYushi(j) = 0
        wdGankin(j) = 0
        wdRisoku(j) = 0
        wdHensai(j) = 0
        wdKaiyaku(j) = 0
        wdYZan(j) = 0
        wdYushi2(j) = 0
        wdGankin2(j) = 0
        wdRisoku2(j) = 0
        wdHensai2(j) = 0
        wdKaiyaku2(j) = 0
        wdYZan2(j) = 0
    Next
'
    wstr = "Select "
    wstr = wstr & "借入番号,"
    For j = 1 To 12
        w番号 = Right("00" + CStr(j), 2)
        
        wstr = wstr & "融資_" & w番号 & ","
        wstr = wstr & "元金_" & w番号 & ","
        wstr = wstr & "利息_" & w番号 & ","
        wstr = wstr & "返済_" & w番号 & ","
        wstr = wstr & "解約_" & w番号 & ","
        wstr = wstr & "残高_" & w番号 & ","
    Next

    wstr = wstr & "融資合計,"
    wstr = wstr & "元金合計,"
    wstr = wstr & "利息合計,"
    wstr = wstr & "返済合計,"
    wstr = wstr & "解約合計,"
    wstr = wstr & "残高合計"
    wstr = wstr & " From DCDA010_借入残高推移表結果_比較"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    Do Until wRs.eof
        For j = 1 To 12
            w番号 = Right("00" + CStr(j), 2)
            
            If wRs("借入番号") = "1現状" Then
                wdYushi(j) = wRs("融資_" & w番号)
                wdGankin(j) = wRs("元金_" & w番号)
                wdRisoku(j) = wRs("利息_" & w番号)
                wdHensai(j) = wRs("返済_" & w番号)
                wdKaiyaku(j) = wRs("解約_" & w番号)
                wdYZan(j) = wRs("残高_" & w番号)
            Else
                wdYushi2(j) = wRs("融資_" & w番号)
                wdGankin2(j) = wRs("元金_" & w番号)
                wdRisoku2(j) = wRs("利息_" & w番号)
                wdHensai2(j) = wRs("返済_" & w番号)
                wdKaiyaku2(j) = wRs("解約_" & w番号)
                wdYZan2(j) = wRs("残高_" & w番号)
            End If
        Next
    
        If wRs("借入番号") = "1現状" Then
            wdYushi(0) = wRs("融資合計")
            wdGankin(0) = wRs("元金合計")
            wdRisoku(0) = wRs("利息合計")
            wdHensai(0) = wRs("返済合計")
            wdKaiyaku(0) = wRs("解約合計")
            wdYZan(0) = wRs("残高合計")
        Else
            wdYushi2(0) = wRs("融資合計")
            wdGankin2(0) = wRs("元金合計")
            wdRisoku2(0) = wRs("利息合計")
            wdHensai2(0) = wRs("返済合計")
            wdKaiyaku2(0) = wRs("解約合計")
            wdYZan2(0) = wRs("残高合計")
        End If
        
        wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
'
    If pHikaku = True Then
        Call MX040_現状借換後対比表
    End If
'
    Exit Sub
'
'----------< ERROR ROUTINE >--------------------------------------------------------
MX040_借入残高推移表比較データ取得_ERR:
    pERR_MES = pPROGRAM_ID + "/ MX040_借入残高推移表比較データ取得() でエラー" + vbCrLf + vbCrLf + _
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
' MX040_借入金返済予定表
'------------------------------------------------
Private Sub MX040_借入金返済予定表(pCsvFileName As String)
'
    Dim ws01 As String
    Dim wsNendo As String
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wstr = ""
    wstr = wstr & "Select "
'    wstr = wstr & "format(M.実際年月日,'" & Gfmt年月日 & "') As 返済年月日,"
    wstr = wstr & "format(M.実際年月日,'" & Gfmtcsv年月日 & "') As 返済年月日,"    '2012.10.23　追加 by k.kunita
    wstr = wstr & "M.借入番号,"
    wstr = wstr & "K.借入内容,"
    wstr = wstr & "KS.借入金種別名,"
    wstr = wstr & "G.銀行番号,"
    wstr = wstr & "G.銀行名,"
    wstr = wstr + "B.部門名,"
    
    wstr = wstr & "IIF(K.営業日区分 = " & P8.FCDbl(XMXA020_区分("営業日", "翌営業日")) & ",'翌営業日','前営業日') As 営業日,"
    wstr = wstr & "IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As 利息区分,"
    wstr = wstr & "IIF(K.利息計算日数区分 = " & P8.FCDbl(XMXA020_区分("利息計算日数", "営業日数")) & ",'営業日数','固定日数') As 利息計算日数,"
    wstr = wstr & "IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As 金利種別,"
    wstr = wstr & "KK.基準金利名,"
    wstr = wstr & "K.金利条件 As 金利備考,"
    
    wstr = wstr & "IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期借入金','長期借入金') AS 長短区分,"
    wstr = wstr & "IIF(有担保フラグ = " & P8.FCDbl(XMXA020_区分("有担フラグ", "無担保")) & ",'無担保','有担保') As 担保区分,"
    wstr = wstr & "K.担保名 As 担保内容,"
    wstr = wstr & "IIF(設備フラグ=1,'設備','運転資金') As 設備区分,"
    wstr = wstr & "K.資金用途,"
    
    wstr = wstr & "format(M.利率,'#,##0.00000') As 利率,"
    wstr = wstr & "M.元金額 As 支払元金額,"
    wstr = wstr & "M.利息額 As 支払利息額,"
    wstr = wstr & "M.返済金額 As 支払額,"
    wstr = wstr & "M.初期手数料+M.元金手数料+M.利息手数料 As 手数料,"
    wstr = wstr & "M.保証料 As 保証料"

    wstr = wstr & " FROM ((((DCDA020_借入金明細 AS M"
    wstr = wstr & " INNER JOIN DBDA010_借入金 AS K"
    wstr = wstr & " ON M.借入番号 = K.借入番号)"
    wstr = wstr & " INNER JOIN DAAA040_銀行マスタ AS G"
    wstr = wstr & " ON K.銀行番号 = G.銀行番号)"
    wstr = wstr + " Inner Join DAAA116_借入金種別 As KS"
    wstr = wstr + " ON K.借入金種別区分 = KS.借入金種別区分)"
    wstr = wstr + " LEFT JOIN DAAA200_部門マスタ AS B"
    wstr = wstr + "  ON B.部門番号 = K.プロジェクト番号)"
    wstr = wstr + " Left Join DAAA116_基準金利 As KK"
    wstr = wstr + " ON K.基準金利区分 = KK.基準金利区分"

    GVar1 = C年月日.平成To西暦("年月", GRpt.テキスト_01)
    GVar2 = C年月日.平成To西暦("年月", GRpt.テキスト_02)

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
    wstr = wstr & " AND KS.社債フラグ=1)"
    
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
    wstr = wstr & " AND KS.社債フラグ=1)"
    
    If GRpt.集計 = "年月日別" Then
        wstr = wstr & " ORDER BY M.実際年月日,K.銀行番号,M.借入番号,M.据置X回目"
    ElseIf GRpt.集計 = "銀行別" Then
        wstr = wstr & " ORDER BY K.銀行番号,M.実際年月日,M.借入番号,M.据置X回目"
    End If
'
    wstr = wstr & " Union Select"
    wstr = wstr & " '総合計','',count(M.借入番号) & '件','','',"
    wstr = wstr & "'','','','','','','','','','','','','','',"
    
    wstr = wstr & "sum(M.元金額),"
    wstr = wstr & "sum(M.利息額),"
    wstr = wstr & "sum(M.返済金額),"
    wstr = wstr & "sum(M.初期手数料)+sum(M.元金手数料)+sum(M.利息手数料) As 手数料,"
    wstr = wstr & "sum(M.保証料)"
    
    wstr = wstr & " FROM ((DCDA020_借入金明細 AS M"
    wstr = wstr & " INNER JOIN DBDA010_借入金 AS K"
    wstr = wstr & " ON M.借入番号 = K.借入番号)"
    wstr = wstr & " INNER JOIN DAAA040_銀行マスタ AS G"
    wstr = wstr & " ON K.銀行番号 = G.銀行番号)"
    wstr = wstr + " Inner Join DAAA116_借入金種別 As KS"
    wstr = wstr + " ON K.借入金種別区分 = KS.借入金種別区分"

    GVar1 = C年月日.平成To西暦("年月", GRpt.テキスト_01)
    GVar2 = C年月日.平成To西暦("年月", GRpt.テキスト_02)

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
    wstr = wstr & " AND KS.社債フラグ=1)"
    
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
    wstr = wstr & " AND KS.社債フラグ=1)"
    
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        '名称
        Write #1, _
            "返済年月日", "借入番号", "借入内容", "借入金種別名", "銀行番号", "銀行名", "部門名", _
            "営業日", "利息区分", "利息計算日数", "金利種別", "基準金利名", "金利備考", _
            "長短区分", "担保区分", "担保内容", "設備区分", "資金用途", _
            "利率", "支払元金額", "支払利息額", "支払額", "手数料", "保証料"
            
    Do Until wRs.eof
        Write #1, _
            P8.FCStr(wRs.Fields("返済年月日").Value), _
            P8.FCStr(wRs.Fields("借入番号").Value), P8.FCStr(wRs.Fields("借入内容").Value), _
            P8.FCStr(wRs.Fields("借入金種別名").Value), _
            P8.FCStr(wRs.Fields("銀行番号").Value), P8.FCStr(wRs.Fields("銀行名").Value), P8.FCStr(wRs.Fields("部門名").Value), _
            P8.FCStr(wRs.Fields("営業日").Value), _
            P8.FCStr(wRs.Fields("利息区分").Value), _
            P8.FCStr(wRs.Fields("利息計算日数").Value), _
            P8.FCStr(wRs.Fields("金利種別").Value), _
            P8.FCStr(wRs.Fields("基準金利名").Value), _
            P8.FCStr(wRs.Fields("金利備考").Value), _
            P8.FCStr(wRs.Fields("長短区分").Value), _
            P8.FCStr(wRs.Fields("担保区分").Value), _
            P8.FCStr(wRs.Fields("担保内容").Value), _
            P8.FCStr(wRs.Fields("設備区分").Value), _
            P8.FCStr(wRs.Fields("資金用途").Value), _
            P8.FCDbl(wRs.Fields("利率").Value), _
            P8.FCDbl(wRs.Fields("支払元金額").Value), _
            P8.FCDbl(wRs.Fields("支払利息額").Value), _
            P8.FCDbl(wRs.Fields("支払額").Value), _
            P8.FCDbl(wRs.Fields("手数料").Value), _
            P8.FCDbl(wRs.Fields("保証料").Value)
    wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_金融機関別残高表
'------------------------------------------------
Public Sub MX040_金融機関別残高表(pCsvFileName As String)
'
    Dim j As Integer
    Dim wdGokei(13) As Double
    Dim wdGZan As Double, wdGTanki As Double, wdGTyoki As Double, wdGShasai As Double
    Dim wdPZan As Double, wdPTanki As Double, wdPTyoki As Double, wdPShasai As Double
    Dim ws01 As String
    Dim wsNendo As String
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    '** レコード　ソース **
    wstr = "Select "
    wstr = wstr & " Sum(W.コード_007) AS 残高合計,"
    wstr = wstr & " Sum(W.コード_008) AS 短期合計,"
    wstr = wstr & " Sum(W.コード_009) AS 長期合計,"
    wstr = wstr & " Sum(W.コード_010) AS 社債合計"
    wstr = wstr & " FROM DCXA020_帳票作成ワーク AS W"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
        wdGZan = P8.FCDbl(wRs("残高合計"))
        wdGTanki = P8.FCDbl(wRs("短期合計"))
        wdGTyoki = P8.FCDbl(wRs("長期合計"))
        wdGShasai = P8.FCDbl(wRs("社債合計"))
    End If
    wRs.Close
    Set wRs = Nothing
    
    wstr = "Select "
    wstr = wstr & "W.科目番号 As 銀行番号,"
    wstr = wstr & "W.科目名 As 銀行名,"
    wstr = wstr & "W.コード_001," 'カウント
    wstr = wstr & "W.コード_002," '融資金額
    wstr = wstr & "W.コード_003," '融資
    wstr = wstr & "W.コード_004," '元金
    wstr = wstr & "W.コード_005," '利息
    wstr = wstr & "W.コード_006," '返済
    wstr = wstr & "W.コード_007," '残高
    wstr = wstr & "W.コード_011," '短期C
    wstr = wstr & "W.コード_008," '短期
    wstr = wstr & "W.コード_012," '長期C
    wstr = wstr & "W.コード_009," '長期
    wstr = wstr & "W.コード_013," '社債C
    wstr = wstr & "W.コード_010"  '社債
    wstr = wstr & " FROM DCXA020_帳票作成ワーク As W"
    wstr = wstr & " Order by W.科目番号"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        '名称
        Write #1, _
            CStr(wRs.Fields("銀行番号").Name), _
            CStr(wRs.Fields("銀行名").Name), _
            "融資金額", _
            "当月融資金額", _
            "元金額", _
            "利息額", _
            "返済金額", _
            "融資残高", _
            "件数", _
            "融資残高構成比", _
            "短期借入金", _
            "短期借入金件数", _
            "短期借入金構成比", _
            "長期借入金", _
            "長期借入金件数", _
            "長期借入金構成比", _
            "社債", _
            "社債件数", _
            "社債構成比"
        
        Do Until wRs.eof
        
            wdPZan = P8.FFix(P8.FCDiv(P8.FCDbl(wRs.Fields("コード_007").Value), wdGZan) * 10000) / 100
            wdPTanki = P8.FFix(P8.FCDiv(P8.FCDbl(wRs.Fields("コード_008").Value), wdGTanki) * 10000) / 100
            wdPTyoki = P8.FFix(P8.FCDiv(P8.FCDbl(wRs.Fields("コード_009").Value), wdGTyoki) * 10000) / 100
            wdPShasai = P8.FFix(P8.FCDiv(P8.FCDbl(wRs.Fields("コード_010").Value), wdGShasai) * 10000) / 100
                    
            Write #1, _
                P8.FCStr(wRs.Fields("銀行番号").Value), _
                P8.FCStr(wRs.Fields("銀行名").Value), _
                P8.FCStr(wRs.Fields("コード_002").Value), _
                P8.FCStr(wRs.Fields("コード_003").Value), _
                P8.FCDbl(wRs.Fields("コード_004").Value), _
                P8.FCDbl(wRs.Fields("コード_005").Value), _
                P8.FCDbl(wRs.Fields("コード_006").Value), _
                P8.FCDbl(wRs.Fields("コード_007").Value), _
                P8.FCStr(wRs.Fields("コード_001").Value), _
                Format(wdPZan, "#,##0.00"), _
                P8.FCDbl(wRs.Fields("コード_008").Value), _
                P8.FCDbl(wRs.Fields("コード_011").Value), _
                Format(wdPTanki, "#,##0.00"), _
                P8.FCDbl(wRs.Fields("コード_009").Value), _
                P8.FCDbl(wRs.Fields("コード_012").Value), _
                Format(wdPTyoki, "#,##0.00"), _
                P8.FCDbl(wRs.Fields("コード_010").Value), _
                P8.FCDbl(wRs.Fields("コード_013").Value), _
                Format(wdPShasai, "#,##0.00")

                For j = 1 To 13
                    ws01 = Right("000" & CStr(j), 3)
                    wdGokei(j) = wdGokei(j) + P8.FCDbl(wRs.Fields("コード_" & ws01).Value)
                Next j
            
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
    
        '合計
        Write #1, _
            "", _
            "合計", _
            wdGokei(2), _
            wdGokei(3), _
            wdGokei(4), _
            wdGokei(5), _
            wdGokei(6), _
            wdGokei(7), _
            wdGokei(1), _
            0, _
            wdGokei(8), _
            wdGokei(11), _
            0, _
            wdGokei(9), _
            wdGokei(12), _
            0, _
            wdGokei(10), _
            wdGokei(13), _
            0
    
    Close #1 '出力ファイルを閉じる
'
'    GDb.Execute wstr
'
'    DoEvents
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_仕訳表
'------------------------------------------------
Private Sub MX040_仕訳表(pCsvFileName As String)
'
    Dim ws01 As String
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wstr = ""
    wstr = wstr & "Select "
    'wstr = wstr & "Format(S.年月日,'" & Gfmt年月日 & "') As 年月日,"
    wstr = wstr & "Format(S.年月日,'" & Gfmtcsv年月日 & "') As 年月日,"     '2012.10.23　追加 by k.kunita
    
    'If GRpt.チェック_02 = 0 Then
    '和暦入力
    '    wstr = wstr & "Format(S.年月日,'ee年mm月dd日') As 年月日,"
   ' Else
    '西暦入力
  '      wstr = wstr & "Format(S.年月日,'yyyy/mm/dd') As 年月日,"
  '  End If
    
    wstr = wstr & "借方勘定科目,"
    wstr = wstr & "借方勘定科目名,"
    wstr = wstr & "借方補助科目,"
    wstr = wstr & "借方補助科目名,"
    wstr = wstr & "借方金額,"
    wstr = wstr & "貸方勘定科目,"
    wstr = wstr & "貸方勘定科目名,"
    wstr = wstr & "貸方補助科目,"
    wstr = wstr & "貸方補助科目名,"
    wstr = wstr & "貸方金額,"
    'wstr = wstr & "仕訳区分,"
    wstr = wstr & "仕訳名,"
    wstr = wstr & "K.銀行番号,"
    wstr = wstr & "G.銀行名,"
    wstr = wstr & "G.預金種別,"
    wstr = wstr & "G.金融機関番号,"
    wstr = wstr & "G.支店番号,"
    wstr = wstr & "G.口座番号,"
    wstr = wstr & "S.借入番号,"
    wstr = wstr & "B.部門番号,"
    wstr = wstr & "B.部門名"
    wstr = wstr & " FROM ((DCDA040_仕訳データ As S"
    wstr = wstr & " LEFT JOIN DBDA010_借入金 As K"
    wstr = wstr & " ON S.借入番号 = K.借入番号)"
    wstr = wstr & " LEFT JOIN DAAA040_銀行マスタ As G"
    wstr = wstr & " ON S.銀行番号 = G.銀行番号)"
    wstr = wstr + " LEFT JOIN DAAA200_部門マスタ AS B"
    wstr = wstr + " ON B.部門番号 = K.プロジェクト番号"
    'wstr = wstr & " Order BY 仕訳区分,仕訳補助,年月日,日番号,G.銀行番号,番号"
    'wstr = wstr & " Order BY Format(年月日,'yyyy/mm'),仕訳区分,仕訳補助,社債フラグ,年月日,S.銀行番号,S.借入番号"
    wstr = wstr & " Order BY 対象年月,仕訳区分,仕訳補助,社債フラグ,年月日,S.銀行番号,S.借入番号"
    
    'If GRpt.帳票名 = "仕訳表 -月次処理-" Then
    '    wstr = wstr & " Order BY S.仕訳区分,S.社債フラグ,仕訳補助,S.年月日,S.日番号,S.銀行番号,S.番号"
    'ElseIf GRpt.帳票名 = "仕訳表 -決算処理-" Then
    '    wstr = wstr & " Order BY Format(S.年月日,'yyyy/mm'),S.仕訳区分,S.社債フラグ,仕訳補助,S.年月日,S.日番号,S.銀行番号"
    'End If
    
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        '名称
        Write #1, _
            "年月日", _
            "借方勘定科目", "借方勘定科目名", "借方補助科目", "借方補助科目名", "借方金額", _
            "貸方勘定科目", "貸方勘定科目名", "貸方補助科目", "貸方補助科目名", "貸方金額", _
            "銀行番号", "銀行名", _
            "金融機関番号", "支店番号", "預金種別", "口座番号", _
            "部門番号", "部門名", "借入番号", "仕訳名"
            
    Do Until wRs.eof
        Write #1, _
            P8.FCStr(wRs.Fields("年月日").Value), _
            P8.FCStr(wRs.Fields("借方勘定科目").Value), _
            P8.FCStr(wRs.Fields("借方勘定科目名").Value), _
            P8.FCStr(wRs.Fields("借方補助科目").Value), _
            P8.FCStr(wRs.Fields("借方補助科目名").Value), _
            P8.FCDbl(wRs.Fields("借方金額").Value), _
            P8.FCStr(wRs.Fields("貸方勘定科目").Value), _
            P8.FCStr(wRs.Fields("貸方勘定科目名").Value), _
            P8.FCStr(wRs.Fields("貸方補助科目").Value), _
            P8.FCStr(wRs.Fields("貸方補助科目名").Value), _
            P8.FCDbl(wRs.Fields("貸方金額").Value), _
            P8.FCStr(wRs.Fields("銀行番号").Value), _
            P8.FCStr(wRs.Fields("銀行名").Value), _
            P8.FCStr(wRs.Fields("金融機関番号").Value), P8.FCStr(wRs.Fields("支店番号").Value), _
            P8.FCStr(wRs.Fields("預金種別").Value), P8.FCStr(wRs.Fields("口座番号").Value), _
            P8.FCStr(wRs.Fields("部門番号").Value), P8.FCStr(wRs.Fields("部門名").Value), P8.FCStr(wRs.Fields("借入番号").Value), P8.FCStr(wRs.Fields("仕訳名").Value)
    wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_勘定科目
'------------------------------------------------
Public Sub MX040_勘定科目(pCsvFileName As String)
'
    Dim ws01 As String
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wstr = ""
    wstr = wstr & "Select "
    wstr = wstr & "仕訳名,"
    wstr = wstr & "IIF(社債フラグ = 0,'','○') As 社債,"
    wstr = wstr & "IIF(仕訳補助備考='長短区分',"
    wstr = wstr & "IIF(仕訳補助='0','長短区分：短期借入金',IIF(仕訳補助='1','長短区分：長期借入金','区分なし')),"
    wstr = wstr & "IIF(仕訳補助備考='利息区分',"
    wstr = wstr & "IIF(仕訳補助='1','利息区分：利息先払',IIF(仕訳補助='2','利息区分：利息後払','区分なし')),"
    wstr = wstr & "'区分なし')) As 仕訳補助名,"
    
    wstr = wstr & "借方勘定科目,"
    wstr = wstr & "借方勘定科目名,"
    wstr = wstr & "IIF(借方補助科目使用 <> 0,'○','×') As 借方補助使用,"
    'wstr = wstr & "IIF(借方個別補助科目使用 <> 0,'○','×') As 借方個別補助使用,"
    wstr = wstr & "貸方勘定科目,"
    wstr = wstr & "貸方勘定科目名,"
    wstr = wstr & "IIF(貸方補助科目使用 <> 0,'○','×') As 貸方補助使用,"
    'wstr = wstr & "IIF(貸方個別補助科目使用 <> 0,'○','×') As 貸方個別補助使用,"
    wstr = wstr & "IIF(取消フラグ = 0,'','×') As 取消"
    wstr = wstr & " From DABA010_勘定科目マスタ"
    wstr = wstr + " Order By 仕訳区分,社債フラグ,仕訳補助"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        '名称
        Write #1, _
            "仕訳名", "社債", "仕訳補助名", _
            "借方勘定科目", "借方勘定科目名", "借方補助使用", _
            "貸方勘定科目", "貸方勘定科目名", "貸方補助使用", _
            "取消"
            
    Do Until wRs.eof
        Write #1, _
            P8.FCStr(wRs.Fields("仕訳名").Value), _
            P8.FCStr(wRs.Fields("社債").Value), _
            P8.FCStr(wRs.Fields("仕訳補助名").Value), _
            P8.FCStr(wRs.Fields("借方勘定科目").Value), _
            P8.FCStr(wRs.Fields("借方勘定科目名").Value), _
            P8.FCStr(wRs.Fields("借方補助使用").Value), _
            P8.FCStr(wRs.Fields("貸方勘定科目").Value), _
            P8.FCStr(wRs.Fields("貸方勘定科目名").Value), _
            P8.FCStr(wRs.Fields("貸方補助使用").Value), _
            P8.FCStr(wRs.Fields("取消").Value)
            
    wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    MsgBox "出力しました", vbInformation
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_補助科目
'------------------------------------------------
Public Sub MX040_補助科目(pCsvFileName As String)
'
    Dim ws01 As String
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wstr = ""
    wstr = wstr & "Select "
    wstr = wstr & "勘定科目,"
    wstr = wstr & "勘定科目名,"
    wstr = wstr & "H.銀行番号,"
    wstr = wstr & "銀行名,"
    wstr = wstr & "補助科目,"
    wstr = wstr & "補助科目名"
    wstr = wstr & " From DABA020_補助科目マスタ As H"
    wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr & " ON H.銀行番号 = G.銀行番号"
    wstr = wstr + " Order By 勘定科目,H.銀行番号,補助科目"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        '名称
        Write #1, _
            "勘定科目", "勘定科目名", "銀行番号", "銀行名", "補助科目", "補助科目名"
            
    Do Until wRs.eof
        Write #1, _
            P8.FCStr(wRs.Fields("勘定科目").Value), _
            P8.FCStr(wRs.Fields("勘定科目名").Value), _
            P8.FCStr(wRs.Fields("銀行番号").Value), _
            P8.FCStr(wRs.Fields("銀行名").Value), _
            P8.FCStr(wRs.Fields("補助科目").Value), _
            P8.FCStr(wRs.Fields("補助科目名").Value)
            
    wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    MsgBox "出力しました", vbInformation
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_個別補助科目
'------------------------------------------------
Public Sub MX040_個別補助科目(pCsvFileName As String)
'
    Dim wRs2 As ADODB.Recordset
    Dim wstr2 As String
    
    Dim ws01 As String
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        '名称
    Write #1, _
        "勘定科目", "勘定科目名", "借入番号", "銀行番号", "銀行名", "個別補助科目", "個別補助科目名"
    wstr2 = ""
    wstr2 = wstr2 & "Select "
    wstr2 = wstr2 & "借方勘定科目 AS 科目, 借方勘定科目名 As 科目名, 借方個別補助科目使用 As 使用"
    wstr2 = wstr2 & " FROM DABA010_勘定科目マスタ"
    wstr2 = wstr2 & " Where 借方個別補助科目使用 = 1"
    wstr2 = wstr2 + " Order By 借方勘定科目"
    wstr2 = wstr2 & " UNION SELECT 貸方勘定科目, 貸方勘定科目名, 貸方個別補助科目使用"
    wstr2 = wstr2 & " From DABA010_勘定科目マスタ"
    wstr2 = wstr2 & " WHERE 貸方個別補助科目使用=1"
    Call AdoRecordsetOpen(GDb, wRs2, wstr2)
    Do Until wRs2.eof
        
        wstr = ""
        wstr = wstr & "Select "
        wstr = wstr & "H.借入番号,"
        wstr = wstr & "H.銀行番号,"
        wstr = wstr & "G.銀行名,"
        wstr = wstr & "H.個別補助科目,"
        wstr = wstr & "H.個別補助科目名"
        wstr = wstr & " FROM DABA030_個別補助科目マスタ As H"
        wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
        wstr = wstr & " ON H.銀行番号 = G.銀行番号"
        wstr = wstr + " Order By H.個別補助科目"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            Write #1, _
                P8.FCStr(wRs2.Fields("科目").Value), _
                P8.FCStr(wRs2.Fields("科目名").Value), _
                P8.FCStr(wRs.Fields("借入番号").Value), _
                P8.FCStr(wRs.Fields("銀行番号").Value), _
                P8.FCStr(wRs.Fields("銀行名").Value), _
                P8.FCStr(wRs.Fields("個別補助科目").Value), _
                P8.FCStr(wRs.Fields("個別補助科目名").Value)
                
        wRs.MoveNext
        Loop
        wRs.Close
        Set wRs = Nothing
    
    wRs2.MoveNext
    Loop
    wRs2.Close
    Set wRs2 = Nothing
    
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    MsgBox "出力しました", vbInformation
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub
    
'------------------------------------------------
' MXA040_COMDLG
'------------------------------------------------
Public Function MXA040_COMDLG(wCDialog As CommonDialog, wTitle As String, wDir As String, wFilter As String, Optional wFile As String = "") As String
'
On Error GoTo ComCancel
    wCDialog.DialogTitle = wTitle
    wCDialog.InitDir = wDir
    wCDialog.Filter = wFilter
    wCDialog.FileName = wFile
    wCDialog.CancelError = True
    
    wCDialog.ShowSave
    MXA040_COMDLG = wCDialog.FileName
'
    Exit Function
'
ComCancel:
    MXA040_COMDLG = "キャンセル"
End Function

'------------------------------------------------
' MXA040_CHECK_FileName
'------------------------------------------------
Public Function MXA040_CHECK_FileName(pPath As String) As Boolean
'
    Dim wi01 As Integer
    Dim j As Integer
    Dim ws01 As String, ws02 As String
'
    On Error GoTo MXA040_CHECK_FileName_ERR
'
    MXA040_CHECK_FileName = False
'
    ws01 = "": ws02 = ""
    wi01 = Len(pPath)
    For j = wi01 To 1 Step -1
        ws01 = Mid$(pPath, j, 1)
        If ws01 <> "\" Then
            ws02 = ws01 & ws02
        Else
            Exit For
        End If
    Next j
'
    If LCase(ws02) <> LCase("借入明細表.csv") Then
        GRet = MsgBox("CSVファイル名が違います。" & vbCrLf & "取り込みますか？", vbYesNo + vbCritical)
        If GRet = vbNo Then
            Exit Function
        End If
    End If
'
    MXA040_CHECK_FileName = True
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MXA040_CHECK_FileName_ERR:
    pERR_MES = pPROGRAM_ID + "/ MXA040_CHECK_FileName() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                "金剛石を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function

'------------------------------------------------
' MXA040_借入明細移行_ワークテーブル作成
'------------------------------------------------
Public Sub MXA040_借入明細移行_ワークテーブル作成(p借入番号 As String)
'
    wstr = "Delete * from DCIA010_借入金ワーク"
    GDb.Execute wstr

    DoEvents
'
    wstr = ""
    wstr = "INSERT INTO DCIA010_借入金ワーク"
    wstr = wstr & " Select * From DBDA010_借入金"
    wstr = wstr & " Where 借入番号 = '" & p借入番号 & "'"
    GDb.Execute wstr
'
End Sub

'------------------------------------------------
' MXA040_借入明細移行
'------------------------------------------------
Public Function MXA040_借入明細移行(p借入番号 As String) As Boolean
'
    '** 明細ファイル 作成 **
    wstr = ""
    wstr = wstr + "Select * From DCIA010_借入金ワーク"
    wstr = wstr + " Where 借入番号 = '" & p借入番号 & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
      
        w借入データ = MBD010_借入データセット(wRs)
        Call MBD010_借入金テーブル作成("", w借入データ)
        Call MXA040_借入明細作成_移行(w借入データ)
        '
        wRs.Update
        
    End If
    wRs.Close
    Set wRs = Nothing
'
End Function

'------------------------------------------------
' MXA040_借入明細作成_移行(MBD010_借入明細作成 作成テーブル変更)
'------------------------------------------------
Public Sub MXA040_借入明細作成_移行(p借入計画マスタ As MAA910_借入金)
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String

    Dim j As Integer
    Dim w解約実行日 As Variant      ' 07/02/21 V180
    Dim w返済回数 As Integer        '10/02/27
'
    On Error GoTo MXA040_借入明細作成_移行_ERR
'
    '** 明細ファイル 削除 **
    wstr = "Delete * From DBDA010_借入金明細TR"
    wstr = wstr & " Where 借入番号='" & p借入計画マスタ.借入番号 & "'"
    GDb.Execute wstr
'
    w返済回数 = 0                   '10/02/27
'
    wstr = ""
    wstr = wstr + "Select * From DBDA010_借入金明細TR"
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
                'wRs("据置X回目") = G借入金テーブル(j).据置X回目
                
                wRs("据置X回目") = 0
                wRs("返済予定年月") = G借入金テーブル(j).返済予定年月
    
                wRs("実際年月日") = G借入金テーブル(j).実際年月日
                wRs("利息計算年月日") = G借入金テーブル(j).利息計算年月日   '10/01/04
                wRs("利息額") = G借入金テーブル(j).利息額
                wRs("保証料") = G借入金テーブル(j).保証料
                wRs("手数料") = G借入金テーブル(j).手数料       ' 08/12/06 V189
                wRs("金融保証料") = G借入金テーブル(j).金融保証料
                
                w解約実行日 = p借入計画マスタ.解約実行日
                
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
                    
                wRs("日割日数") = G借入金テーブル(j).日割日数
                wRs("利率") = G借入金テーブル(j).利率
                
                If p借入計画マスタ.日割計算区分 = CDbl(XMXA020_区分("日割計算区分", "自動計算")) Then
                    wRs("利息対象期間日数") = 0
                Else
                    wRs("利息対象期間日数") = G借入金テーブル(j).利息対象期間日数       'V182 2008/01/28
                End If

                
            wRs.Update
          End If                    '10/02/27
          
        Next
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MXA040_借入明細作成_移行_ERR:
    pERR_MES = pPROGRAM_ID + "/ MXA040_借入明細作成_移行() でエラー" + vbCrLf + vbCrLf + _
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
' MXA040_借入明細取込
'------------------------------------------------
Public Function MXA040_借入明細取込(pCsvName As String, pBango As String) As Boolean
'
    Dim j As Long, k As Long, l As Long
    Dim wi01 As Integer, wiKcnt As Integer, wiScnt As Integer
    Dim ws01 As String
    Dim wsName As String, wsValue As String
    Dim wsMsg As String
    Dim GFLG_RisokuKeisan() As Boolean
    Dim wDate1 As Date, wDate2 As Date, wDate3 As Date
'
    MXA040_借入明細取込 = False
'
    On Error GoTo MXA040_借入明細取込_ERR
'
    w借入データ.社債フラグ = GInt1
    '項目名で取り込むので名称セット
    Call MXA040_借入明細項目名_SET(pCsvName)
'
    ReDim GFLG_RisokuKeisan(UBound(wレコード))
    ReDim G借入金入力(0)
    ReDim w社債入力(0)

    If UBound(wレコード) = 0 Then
        Exit Function
    End If
'
    wiKcnt = 0
    wiScnt = 0
    For k = 1 To UBound(wレコード)
        
        GFLG_RisokuKeisan(k) = False
        
        For j = 1 To UBound(wレコード(k).xName)
            wsName = wレコード(k).xName(j)
            wsValue = CStr(wレコード(k).xValue(j))
            
            'Data Check
            If wsName = "借入番号" Then
                If wsValue = "" Then
                    GoTo MXA040_借入明細取込_ERR_CHECK
                End If
                If pBango <> wsValue And pBango <> "" Then
                    GoTo MXA040_借入明細取込_ERR_CHECK
                End If
            ElseIf wsName = "返済年月日" Then
                ws01 = Format(wsValue, "yyyy/mm/dd")
                If Not IsDate(ws01) Then
                    GoTo MXA040_借入明細取込_ERR_CHECK
                End If
            ElseIf wsName = "利息計算年月日" Then
                ws01 = Format(wsValue, "yyyy/mm/dd")
                If Not IsDate(ws01) Then
                    GoTo MXA040_借入明細取込_ERR_CHECK
                End If
            ElseIf wsName = "利息額" Then
                If wsValue = "" Then
                    GFLG_RisokuKeisan(k) = True
                End If
                If Not IsNumeric(wsValue) And wsValue <> "" Then
                    GoTo MXA040_借入明細取込_ERR_CHECK
                End If
            
            ElseIf wsName = "初期手数料" Then
                If Not IsNumeric(wsValue) And wsValue <> "" Then
                    GoTo MXA040_借入明細取込_ERR_CHECK
                End If
            ElseIf wsName = "元金手数料" Then
                If Not IsNumeric(wsValue) And wsValue <> "" Then
                    GoTo MXA040_借入明細取込_ERR_CHECK
                End If
            ElseIf wsName = "利息手数料" Then
                If Not IsNumeric(wsValue) And wsValue <> "" Then
                    GoTo MXA040_借入明細取込_ERR_CHECK
                End If
            ElseIf wsName = "保証料" Then
                If Not IsNumeric(wsValue) And wsValue <> "" Then
                    GoTo MXA040_借入明細取込_ERR_CHECK
                End If
            
            Else
                If Not IsNumeric(wsValue) And wsValue <> "" Then
                    GoTo MXA040_借入明細取込_ERR_CHECK
                End If
            End If
        
        Next j
    
        If P8.FCDbl(wレコード(k).xValue(5)) <> 0 _
        Or P8.FCDbl(wレコード(k).xValue(6)) <> 0 Then
            wiKcnt = wiKcnt + 1
            ReDim Preserve G借入金入力(wiKcnt)
            
            G借入金入力(wiKcnt).借入返済年月日 = Format(wレコード(k).xValue(3), "yyyy/mm/dd")
            G借入金入力(wiKcnt).利息計算年月日 = Format(wレコード(k).xValue(4), "yyyy/mm/dd")
            G借入金入力(wiKcnt).元金 = P8.FCDbl(wレコード(k).xValue(5))
            G借入金入力(wiKcnt).利息額 = P8.FCDbl(wレコード(k).xValue(6))
            G借入金入力(wiKcnt).仮計上利息額 = P8.FCDbl(wレコード(k).xValue(7))
            G借入金入力(wiKcnt).返済金額 = P8.FCDbl(wレコード(k).xValue(8))
            G借入金入力(wiKcnt).融資残高 = P8.FCDbl(wレコード(k).xValue(9))
            G借入金入力(wiKcnt).日割日数 = P8.FCDbl(wレコード(k).xValue(10))
            G借入金入力(wiKcnt).利息対象期間日数 = P8.FCDbl(wレコード(k).xValue(11))
            G借入金入力(wiKcnt).利率 = P8.FCDbl(wレコード(k).xValue(12))
        End If
        
        If w借入データ.社債フラグ = 1 Then
            If P8.FCDbl(wレコード(k).xValue(13)) <> 0 _
            Or P8.FCDbl(wレコード(k).xValue(14)) <> 0 _
            Or P8.FCDbl(wレコード(k).xValue(15)) <> 0 _
            Or P8.FCDbl(wレコード(k).xValue(17)) <> 0 Then
                wiScnt = wiScnt + 1
                ReDim Preserve w社債入力(wiScnt)
                
                w社債入力(wiScnt).借入返済年月日 = Format(wレコード(k).xValue(3), "yyyy/mm/dd")
                w社債入力(wiScnt).初期手数料 = P8.FCDbl(wレコード(k).xValue(13))
                w社債入力(wiScnt).元金手数料 = P8.FCDbl(wレコード(k).xValue(14))
                w社債入力(wiScnt).利息手数料 = P8.FCDbl(wレコード(k).xValue(15))
                w社債入力(wiScnt).保証料 = P8.FCDbl(wレコード(k).xValue(17))
            End If
        End If
    Next k
'
    ' =========================================
    '         借入金データセット
    ' =========================================
    'csv借入番号
    ws01 = CStr(wレコード(1).xValue(1))
    
    '借入金データセット
    wstr = ""
    wstr = wstr & "SELECT * From DBDA010_借入金"
    wstr = wstr & " Where 借入番号 ='" & ws01 & "'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
        w借入データ = MBD010_借入データセット(wRs)
    End If
    wRs.Close
    Set wRs = Nothing
'
    ' =========================================
    '         借入金データとのCHECK
    ' =========================================
    '借入番号とのCHECK
    If w借入データ.借入番号 <> ws01 Then
        MsgBox "借入番号:" & ws01 & "は登録されていません。", vbInformation
        
        Exit Function
    End If

    '入力登録のCHECK
    If P8.FCDbl(w借入データ.手入力区分) = P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
        MsgBox "登録方法:標準登録のデータはCSV明細データを取り込めません。", vbInformation
        
        Exit Function
    End If
'
    ' =========================================
    '               CSVデータCHECK
    ' =========================================
    '1行目のCHECK
    k = 1
    
    '利息先払のデータCHECK
    If w借入データ.社債フラグ = 0 Then
        If P8.FCDbl(w借入データ.利息区分) = P8.FCDbl(XMXA020_区分("利息区分", "利息先払")) Then
        '1行目は借入返済年月日=実行日
            If Format(w借入データ.実行日, "yyyy/mm/dd") <> Format(G借入金入力(k).借入返済年月日, "yyyy/mm/dd") Then
                wsMsg = CStr(k) & "行目 借入返済年月日:" & G借入金入力(k).借入返済年月日 & " を確認してください。"
                MsgBox wsMsg, vbInformation
            
                Exit Function
            End If
        End If
    End If
    '
        
    '実行日とのCHECK
    If Format(w借入データ.実行日, "yyyy/mm/dd") > Format(G借入金入力(k).借入返済年月日, "yyyy/mm/dd") Then
        wsMsg = CStr(k) & "行目 借入返済年月日:" & G借入金入力(k).借入返済年月日 & " を確認してください。"
        MsgBox wsMsg, vbInformation
    
        Exit Function
    End If
    
    If Format(w借入データ.実行日, "yyyy/mm/dd") > Format(G借入金入力(k).利息計算年月日, "yyyy/mm/dd") Then
        wsMsg = CStr(k) & "行目 利息計算年月日:" & G借入金入力(k).利息計算年月日 & " を確認してください。"
        MsgBox wsMsg, vbInformation
    
        Exit Function
    End If
    
    '利息後払のデータCHECK
    If P8.FCDbl(w借入データ.利息区分) = P8.FCDbl(XMXA020_区分("利息区分", "利息後払")) Then
    '実行日のデータがあれば利息額=0
        If Format(w借入データ.実行日, "yyyy/mm/dd") = Format(G借入金入力(k).借入返済年月日, "yyyy/mm/dd") Then
            If G借入金入力(k).利息額 <> 0 Then
                wsMsg = CStr(k) & "行目 利息額:" & G借入金入力(k).利息額 & " を確認してください。" + vbCrLf & vbCrLf & _
                        "利息後払の場合、実行日の利息額は0になります。"
                MsgBox wsMsg, vbInformation
            
                Exit Function
            End If
        End If
    End If
    
    '利息先払のデータCHECK
    If P8.FCDbl(w借入データ.利息区分) = P8.FCDbl(XMXA020_区分("利息区分", "利息先払")) Then
    '実行日のデータがあれば利息額<>0のMsg
        If Format(w借入データ.実行日, "yyyy/mm/dd") = Format(G借入金入力(k).借入返済年月日, "yyyy/mm/dd") Then
            If G借入金入力(k).利息額 = 0 Then
                wsMsg = CStr(k) & "行目 利息額:" & G借入金入力(k).利息額 & " を確認してください。" + vbCrLf & vbCrLf & _
                        "実行日の利息額=0で登録しますか？。"
                GRet = MsgBox(wsMsg, vbYesNo + vbQuestion)
                If GRet <> vbYes Then
                    Exit Function
                End If
            End If
        End If
    End If
    
    For k = 2 To UBound(G借入金入力)
        '実行日とのCHECK
        If Format(w借入データ.実行日, "yyyy/mm/dd") > Format(G借入金入力(k).借入返済年月日, "yyyy/mm/dd") Then
            wsMsg = CStr(k) & "行目 借入返済年月日:" & G借入金入力(k).借入返済年月日 & " を確認してください。"
            MsgBox wsMsg, vbInformation
        
            Exit Function
        End If
        
        If Format(w借入データ.実行日, "yyyy/mm/dd") > Format(G借入金入力(k).利息計算年月日, "yyyy/mm/dd") Then
            wsMsg = CStr(k) & "行目 利息計算年月日:" & G借入金入力(k).利息計算年月日 & " を確認してください。"
            MsgBox wsMsg, vbInformation
        
            Exit Function
        End If
        
        '昇順CHECK
        If Format(G借入金入力(k).借入返済年月日, "yyyy/mm/dd") <= Format(G借入金入力(k - 1).借入返済年月日, "yyyy/mm/dd") Then
            wsMsg = CStr(k) & "行目 借入返済年月日:" & G借入金入力(k).借入返済年月日 & " を確認してください。"
            MsgBox wsMsg, vbInformation
        
            Exit Function
        End If
        
        If Format(G借入金入力(k).利息計算年月日, "yyyy/mm/dd") <= Format(G借入金入力(k - 1).利息計算年月日, "yyyy/mm/dd") Then
            wsMsg = CStr(k) & "行目 利息計算年月日:" & G借入金入力(k).利息計算年月日 & " を確認してください。"
            MsgBox wsMsg, vbInformation
        
            Exit Function
        End If
    
    Next k
'
    w借入データ.初回返済実行日 = G借入金入力(1).借入返済年月日
    w借入データ.最終返済実行日 = G借入金入力(UBound(G借入金入力)).借入返済年月日
    
    '利息日数、利息額、返済額の再計算
    If w借入データ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
        wDate2 = Format(w借入データ.最終返済実行日, "yyyy/mm/dd")
    
        For k = UBound(G借入金入力) To 1 Step -1
            
            If Format(G借入金入力(k).借入返済年月日, "yyyy/mm/dd") < Format(w借入データ.最終返済実行日, "yyyy/mm/dd") Then
                
                wDate3 = Format(G借入金入力(k).借入返済年月日, "yyyy/mm/dd")
                
                wDate1 = Format(G借入金入力(k).利息計算年月日, "yyyy/mm/dd")
                wi01 = DateDiff("d", wDate1, wDate2)
                
                If Format(wDate3, "yyyy/mm/dd") = Format(w借入データ.実行日, "yyyy/mm/dd") Then
                'wdate1が実行日の場合
                    
                    '実行日を含めた日数
                    wi01 = wi01 + 1
                    
                    '実行日控除
                    If w借入データ.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日控除")) _
                    Or w借入データ.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) Then
                        wi01 = wi01 - 1
                    End If
                End If
                
                If Format(wDate2, "yyyy/mm/dd") = Format(w借入データ.最終返済実行日, "yyyy/mm/dd") Then
                '最終返済日より前の1件目
                    
                    '最終返済実行日を除く日数
                    'wi01 = wi01 - 1
                    
                    '最終返済日控除
                    If w借入データ.利息控除区分 = CDbl(XMXA020_区分("利息控除", "最終返済日控除")) _
                    Or w借入データ.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) Then
                        wi01 = wi01 - 1
                    End If
                End If
                
                G借入金入力(k).日割日数 = wi01
                
                '利息再計算
                If GFLG_RisokuKeisan(k) = True Then
                    G借入金入力(k).利息額 = MBD010_利息計算小数点5桁(G借入金入力(k).利率, _
                                G借入金入力(k).融資残高, G借入金入力(k).日割日数, w借入データ.金利計算年間日数)
                    G借入金入力(k).返済金額 = G借入金入力(k).利息額 + G借入金入力(k).元金
                End If
                
                wDate2 = wDate1
                
            End If
        
        Next k
        
    Else
    
        wDate1 = Format(w借入データ.実行日, "yyyy/mm/dd")
    
        For k = 1 To UBound(G借入金入力)
            If Format(G借入金入力(k).借入返済年月日, "yyyy/mm/dd") > Format(w借入データ.実行日, "yyyy/mm/dd") Then
                
                wDate3 = Format(G借入金入力(k).借入返済年月日, "yyyy/mm/dd")
                
                wDate2 = Format(G借入金入力(k).利息計算年月日, "yyyy/mm/dd")
                wi01 = DateDiff("d", wDate1, wDate2)
                
                If Format(wDate1, "yyyy/mm/dd") = Format(w借入データ.実行日, "yyyy/mm/dd") Then
                '実行日移行の1件目
                    
                    '実行日を含めた日数
                    wi01 = wi01 + 1
                    
                    '実行日控除
                    If w借入データ.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日控除")) _
                    Or w借入データ.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) Then
                        wi01 = wi01 - 1
                    End If
                End If
                
                If Format(wDate3, "yyyy/mm/dd") = Format(w借入データ.最終返済実行日, "yyyy/mm/dd") Then
                'wdate1が最終返済日の場合
                    
                    '最終返済日控除
                    If w借入データ.利息控除区分 = CDbl(XMXA020_区分("利息控除", "最終返済日控除")) _
                    Or w借入データ.利息控除区分 = CDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) Then
                        wi01 = wi01 - 1
                    End If
                End If
                
                G借入金入力(k).日割日数 = wi01
                
                '利息再計算
                If GFLG_RisokuKeisan(k) = True Then
                    G借入金入力(k).利息額 = MBD010_利息計算小数点5桁(G借入金入力(k).利率, _
                                G借入金入力(k).融資残高 + G借入金入力(k).元金, G借入金入力(k).日割日数, w借入データ.金利計算年間日数)
                    G借入金入力(k).返済金額 = G借入金入力(k).利息額 + G借入金入力(k).元金
                End If
                
                wDate1 = wDate2
                
            End If
        Next k
    End If
'
    ' =========================================
    '               借入明細UPDATE
    ' =========================================
    If UBound(G借入金入力) >= 1 Then
        wstr = "Delete * From DBDA010_借入金明細TR"
        wstr = wstr & " Where 借入番号='" & w借入データ.借入番号 & "'"
        GDb.Execute wstr
    End If
    '
    DoEvents
    '
    Call MXA040_借入明細UPDATE(ws01)
'
    ' =========================================
    '            借入金UPDATE
    ' =========================================
    wstr = ""
    wstr = wstr & "Update DBDA010_借入金"
    
    If P8.FCDbl(G借入金入力(UBound(G借入金入力)).融資残高) = 0 Then
        wstr = wstr & " Set 手入力区分=1,"
    Else
        wstr = wstr & " Set 手入力区分=2,"
    End If
    
    wstr = wstr & "初回返済年月 =#" & Format(P8.FCDate(w借入データ.初回返済実行日), "yyyy/mm/dd") & "#,"
    wstr = wstr & "初回返済実行日 =#" & Format(P8.FCDate(w借入データ.初回返済実行日), "yyyy/mm/dd") & "#,"
    wstr = wstr & "最終返済年月 =#" & Format(P8.FCDate(w借入データ.最終返済実行日), "yyyy/mm/dd") & "#,"
    wstr = wstr & "最終返済実行日 =#" & Format(P8.FCDate(w借入データ.最終返済実行日), "yyyy/mm/dd") & "#"
    wstr = wstr & " Where 借入番号='" & w借入データ.借入番号 & "'"
    GDb.Execute wstr
'
    ' =========================================
    '               社債明細UPDATE
    ' =========================================
    If UBound(w社債入力) >= 1 Then
        wstr = "Delete * From DBDA010_借入金明細TR2"
        wstr = wstr & " Where 借入番号='" & w借入データ.借入番号 & "'"
        GDb.Execute wstr
    End If
    '
    DoEvents
    '
    Call MXA040_社債明細UPDATE(ws01)
'
    ReDim GFLG_RisokuKeisan(0)
    ReDim G借入金入力(0)
    ReDim w社債入力(0)

     GStr_1 = w借入データ.借入番号
'
    On Error GoTo 0
'
    MXA040_借入明細取込 = True
'
Exit Function
'----------< ERROR >----------------------------------------------------------------
MXA040_借入明細取込_ERR_CHECK:
    wsMsg = CStr(k) & "行目 " & wsName & ":" & wsValue & " を確認してください。"
    MsgBox wsMsg, vbInformation
    
    Exit Function
'
MXA040_借入明細取込_ERR:
    Err.Clear
    Exit Function
End Function

'------------------------------------------------
' MXA040_借入明細UPDATE
'------------------------------------------------
Public Sub MXA040_借入明細UPDATE(p借入番号 As String)
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String

    Dim j As Integer
'
    On Error GoTo MXA040_借入明細UPDATE_ERR
'
    wstr = ""
    wstr = wstr + "Select * From DBDA010_借入金明細TR"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        For j = 1 To UBound(G借入金入力)
             
            wRs.AddNew
        
                wRs("借入番号") = p借入番号
                wRs("返済回数") = j
                wRs("返済予定年月") = G借入金入力(j).借入返済年月日
                wRs("実際年月日") = G借入金入力(j).借入返済年月日
                wRs("利息計算年月日") = G借入金入力(j).利息計算年月日
                
                wRs("元金額") = G借入金入力(j).元金
                wRs("利息額") = G借入金入力(j).利息額
                wRs("返済金額") = G借入金入力(j).返済金額
                wRs("融資残高") = G借入金入力(j).融資残高
                wRs("仮計上利息額") = G借入金入力(j).仮計上利息額
                
                wRs("日割日数") = G借入金入力(j).日割日数
                wRs("利息対象期間日数") = G借入金入力(j).利息対象期間日数
                wRs("利率") = G借入金入力(j).利率
                
                wRs("据置X回目") = 0
                wRs("保証料") = 0
                wRs("金融保証料") = 0
            
            wRs.Update
          
        Next
    
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MXA040_借入明細UPDATE_ERR:
    pERR_MES = pPROGRAM_ID + "/ MXA040_借入明細UPDATE() でエラー" + vbCrLf + vbCrLf + _
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
' MXA040_社債明細UPDATE
'------------------------------------------------
Public Sub MXA040_社債明細UPDATE(p借入番号 As String)
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String

    Dim j As Integer
'
    On Error GoTo MXA040_社債明細UPDATE_ERR
'
    wstr = ""
    wstr = wstr + "Select * From DBDA010_借入金明細TR2"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        For j = 1 To UBound(w社債入力)
             
            wRs.AddNew
        
                wRs("借入番号") = p借入番号
                wRs("返済予定年月") = CDate(Format(w社債入力(j).借入返済年月日, "yyyy/mm") & "/01")
                wRs("実際年月日") = w社債入力(j).借入返済年月日
                
                wRs("初期手数料") = w社債入力(j).初期手数料
                wRs("元金手数料") = w社債入力(j).元金手数料
                wRs("利息手数料") = w社債入力(j).利息手数料
                wRs("保証料") = w社債入力(j).保証料
            
            wRs.Update
          
        Next
    
    wRs.Close
    Set wRs = Nothing
'
    wstr = "Delete * From DBDA010_借入金明細TR2"
    wstr = wstr & " Where 保証料=0 And 初期手数料=0 And 元金手数料=0 And 利息手数料=0"
    GDb.Execute wstr
'
    DoEvents
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MXA040_社債明細UPDATE_ERR:
    pERR_MES = pPROGRAM_ID + "/ MXA040_社債明細UPDATE() でエラー" + vbCrLf + vbCrLf + _
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
' MXA040_借入明細項目名_SET
'------------------------------------------------
Private Sub MXA040_借入明細項目名_SET(pCsvName As String)
'
    '----------< 読込 >-------------------------------------------------------------
    Call MXA040_CsvInit
    '
    Call MXA040_CsvAdd("借入番号")
    Call MXA040_CsvAdd("返済回数")
    Call MXA040_CsvAdd("返済年月日", "d")
    Call MXA040_CsvAdd("利息計算年月日", "d")
    Call MXA040_CsvAdd("元金額")
    Call MXA040_CsvAdd("利息額")
    Call MXA040_CsvAdd("調整利息額")
    Call MXA040_CsvAdd("返済金額")
    Call MXA040_CsvAdd("融資残高")
    Call MXA040_CsvAdd("日割日数")
    
    Call MXA040_CsvAdd("利息対象期間日数")
    Call MXA040_CsvAdd("利率")
    
    If w借入データ.社債フラグ = 1 Then
        Call MXA040_CsvAdd("初期手数料")
        Call MXA040_CsvAdd("元金手数料")
        Call MXA040_CsvAdd("利息手数料")
        Call MXA040_CsvAdd("手数料計")
        Call MXA040_CsvAdd("保証料")
        Call MXA040_CsvAdd("支払計")
    End If
'
    wレコード = MXA040_CsvRead(pCsvName, 2) '1行目タイトル
'
End Sub

'------------------------------------------------
' MXA040_JisekiPath
'------------------------------------------------
Public Function MXA040_JisekiPath() As String
'
    Dim objFileSystem As Object
    Dim objFile As Object
    Dim ws01 As String
'
    MXA040_JisekiPath = ""
'
    ws01 = GCurDir & "\" & GTemp
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFileSystem.GetFile(ws01)
        ws01 = UCase(objFile.Drive)
    Set objFile = Nothing
    Set objFileSystem = Nothing
    
    MXA040_JisekiPath = ws01 & "\" & GJiseki_DirName
'
End Function

'------------------------------------------------
' MXA040_CsvInit
'------------------------------------------------
Public Sub MXA040_CsvInit()
    ReDim wField(0)
End Sub

'------------------------------------------------
' MXA040_CsvAdd
'------------------------------------------------
Public Sub MXA040_CsvAdd(pFieldName As String, Optional pType As String = "S")
'
    Dim j As Integer
'
    j = UBound(wField) + 1
    ReDim Preserve wField(j)
    
    wField(j).Name = Trim(pFieldName)
    
    Select Case LCase(Trim(pType))
        Case "s", "str", "string":  wField(j).Type = "S"
        Case "n", "num", "numeric": wField(j).Type = "N"
        Case "d", "dat", "date":    wField(j).Type = "D"
        Case "f", "flg", "flag":    wField(j).Type = "F"
    End Select
End Sub

'------------------------------------------------
' MXA040_CsvRead
'------------------------------------------------
Public Function MXA040_CsvRead(pCsvName As String, Optional pStartRec As Integer = 1) As MGG010_Typeレコード()
'
    Dim j As Long, wRecNo As Long
    Dim k As Integer, wFileNo As Integer
    Dim wString As String, wData As String
    Dim wRecord() As MGG010_Typeレコード

    Dim ws01 As String
'
    On Error GoTo MXA040_CsvRead_ERR
'
    wFileNo = FreeFile
    ReDim wRecord(0)
    
    Open pCsvName For Input As wFileNo Len = 32000
        For j = 1 To pStartRec - 1
            If eof(wFileNo) Then
                Exit For
            End If
            
            Line Input #wFileNo, wString
        Next
    
        While Not eof(wFileNo)
        
            wRecNo = UBound(wRecord) + 1
            
            ReDim Preserve wRecord(wRecNo)
            ReDim Preserve wRecord(wRecNo).xName(UBound(wField))
            ReDim Preserve wRecord(wRecNo).xValue(UBound(wField))
        
            For j = 1 To UBound(wField)
                wRecord(wRecNo).xName(j) = wField(j).Name
            Next
         
            Line Input #wFileNo, wString
            
            j = 0
            'k = InStr(wString, """,""")
            k = InStr(wString, ",")
            Do While (k <> 0)
                j = j + 1
                
                If j > UBound(wField) Then
                    Exit Do
                End If
                
                'wData = Left(wString, k)
                wData = Left(wString, k - 1)
                
                If Left(Trim(wData), 1) = """" _
                And Right(Trim(wData), 1) = """" Then
                    wData = Trim(wData)
                    wData = Mid$(wData, 2, Len(wData) - 2)
                End If
                
                Select Case wField(j).Type
                    Case "S": wRecord(wRecNo).xValue(j) = wData
                    Case "N": wRecord(wRecNo).xValue(j) = P8.FCDbl(wData)
                    Case "D"
                        If InStr(wData, "/") = 0 Then
                            ws01 = wData
                            ws01 = Replace(ws01, "-", "")
                            ws01 = Replace(ws01, ".", "")
                            wData = Format(ws01, "0000/00/00")
                        End If
                        If Not IsDate(wData) Then
                        'If Not IsDate(Format(P8.FCStr(wData), "yyyy/mm/dd")) Then
                            wRecord(wRecNo).xValue(j) = wData
'                            Close wFileNo
'                            Exit Function
                        Else
                            wRecord(wRecNo).xValue(j) = P8.FCDate(wData)
                        End If
                    Case "F":
                        Select Case LCase(Trim(wData))
                            Case "0", "false", ""
                                wRecord(wRecNo).xValue(j) = 0
                            Case Else
                                wRecord(wRecNo).xValue(j) = 1
                        End Select
                End Select
                
                wString = Right(wString, Len(wString) - k)
                'wString = Right(wString, Len(wString) - k - 1)
            
                k = InStr(wString, ",")
                'k = InStr(wString, """,""")
            Loop
            '
            If j < UBound(wField) Then
                j = UBound(wField)
                wData = wString

                If Left(Trim(wData), 1) = """" _
                And Right(Trim(wData), 1) = """" Then
                    wData = Trim(wData)
                    wData = Mid$(wData, 2, Len(wData) - 2)
                End If
                                
                Select Case wField(j).Type
                    Case "S": wRecord(wRecNo).xValue(j) = wData
                    Case "N": wRecord(wRecNo).xValue(j) = P8.FCDbl(wData)
                    Case "D": wRecord(wRecNo).xValue(j) = P8.FCDate(wData)
                    Case "F":
                        Select Case LCase(Trim(wData))
                            Case "0", "false", ""
                                wRecord(wRecNo).xValue(j) = 0
                            Case Else
                                wRecord(wRecNo).xValue(j) = 1
                        End Select
                End Select
            End If
            '
        Wend
    
    Close wFileNo
'
    MXA040_CsvRead = wRecord
'
    On Error GoTo 0
'
Exit Function
MXA040_CsvRead_ERR:
    Err.Clear
    Exit Function
End Function

'------------------------------------------------
' MX040_仕訳表_神姫バス
'------------------------------------------------
Private Sub MX040_仕訳表_神姫バス(pCsvFileName As String)
'
    Dim ws01 As String
    Dim strKashi As String
    Dim strKari As String
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wstr = ""
    wstr = wstr & "Select "
    wstr = wstr & "'0001' As 番号,"
    wstr = wstr & "Format(S.年月日,'" & Gfmtcsv年月日 & "') As 年月日,"     '2012.10.23　追加 by k.kunita
    wstr = wstr & "S.伝票番号,"
    wstr = wstr & "S.借方勘定科目,"
    wstr = wstr & "S.借方補助科目,"
    wstr = wstr & "S.貸方勘定科目,"
    wstr = wstr & "S.貸方補助科目,"
    wstr = wstr & "S.借方金額,"
    wstr = wstr & "S.摘要"
    wstr = wstr & " FROM ((DCDA040_仕訳データ2 As S"
    wstr = wstr & " LEFT JOIN DBDA010_借入金 As K"
    wstr = wstr & " ON S.借入番号 = K.借入番号)"
    wstr = wstr & " LEFT JOIN DAAA040_銀行マスタ As G"
    wstr = wstr & " ON S.銀行番号 = G.銀行番号)"
    wstr = wstr + " LEFT JOIN DAAA200_部門マスタ AS B"
    wstr = wstr + " ON B.部門番号 = K.プロジェクト番号"
    wstr = wstr & " Order BY 対象年月,仕訳区分,年月日,S.銀行番号,S.借入番号"
    
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
            
    Do Until wRs.eof
        
        If P8.FCStr(wRs.Fields("借方補助科目").Value) = "" Then
            strKari = P8.FCStr(wRs.Fields("借方勘定科目").Value)
        Else
            strKari = P8.FCStr(wRs.Fields("借方補助科目").Value)
        End If
        
        If P8.FCStr(wRs.Fields("貸方補助科目").Value) = "" Then
            strKashi = P8.FCStr(wRs.Fields("貸方勘定科目").Value)
        Else
            strKashi = P8.FCStr(wRs.Fields("貸方補助科目").Value)
        End If
    
        Write #1, _
            P8.FCStr(wRs.Fields("番号").Value), _
            P8.FCStr(wRs.Fields("年月日").Value), _
            "", _
            P8.FCStr(wRs.Fields("伝票番号").Value), _
            "", _
            strKari, _
            "", _
            "", _
            strKashi, _
            "", _
            P8.FCDbl(wRs.Fields("借方金額").Value), _
            P8.FCStr(wRs.Fields("摘要").Value)
        wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_長短振替表_神姫バス
'------------------------------------------------
Private Sub MX040_長短振替表_神姫バス(pCsvFileName As String)
'
    Dim ws01 As String
    Dim strKashi As String
    Dim strKari As String
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wstr = ""
    wstr = wstr & "Select"
    wstr = wstr & " G.銀行名 As 銀行名,"
    wstr = wstr & " IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期','長期') As 長短区分,"
    wstr = wstr & " IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動','固定') As 金利種別,"
    wstr = wstr & " M.借入番号,"
    wstr = wstr & " M.融資残高 As 基準日時点残高,"
    wstr = wstr & " M.元金額 As 長短振替額,"
    wstr = wstr & " format(K.最終返済実行日,'" & Gfmt年月日 & "') As 終了日"
    
    wstr = wstr & " FROM ((DCKA010_資金繰表 AS M"
    wstr = wstr & " INNER JOIN DBDA010_借入金 AS K"
    wstr = wstr & " ON M.借入番号 = K.借入番号)"
    wstr = wstr & " INNER JOIN DAAA040_銀行マスタ AS G"
    wstr = wstr & " ON K.銀行番号 = G.銀行番号)"
    wstr = wstr & " INNER JOIN DAAA116_借入金種別 AS S"
    wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分"
    
    wstr = wstr & " WHERE S.借入金種別区分='01'" '借入金のみ
    If GRpt.指定 <> "" Then
        wstr = wstr & " And G.銀行名='" & GRpt.指定 & "'"
    End If
    
    wstr = wstr & " ORDER BY K.銀行番号,K.長短区分 desc,K.金利種別 desc,M.借入番号"
    
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
            
        '名称
        Write #1, _
            "銀行名", "長短区分", "金利種別", "借入番号", "基準日時点残高", "長短振替額", "終了日"
            
    Do Until wRs.eof
        
        Write #1, _
            P8.FCStr(wRs.Fields("銀行名").Value), _
            P8.FCStr(wRs.Fields("長短区分").Value), _
            P8.FCStr(wRs.Fields("金利種別").Value), _
            P8.FCStr(wRs.Fields("借入番号").Value), _
            P8.FCDbl(wRs.Fields("基準日時点残高").Value), _
            P8.FCDbl(wRs.Fields("長短振替額").Value), _
            P8.FCStr(wRs.Fields("終了日").Value)
            
        wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_資金繰表_神姫バス
'------------------------------------------------
Private Sub MX040_資金繰表_神姫バス(pCsvFileName As String)
'
    Dim ws01 As String
    Dim strKashi As String
    Dim strKari As String
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wstr = ""
    wstr = wstr & "Select"
    wstr = wstr & " G.銀行名 As 銀行名,"
    wstr = wstr & " IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期','長期') As 長短区分,"
    wstr = wstr & " IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動','固定') As 金利種別,"
    wstr = wstr & " M.借入番号 As 借入番号,"
    wstr = wstr & " format(M.実際年月日,'" & Gfmt年月日 & "') As 支払日,"
    wstr = wstr & " M.融資金額 As 融資金額,"
    wstr = wstr & " M.元金額 As 元金額,"
    wstr = wstr & " M.利息額 As 利息額,"
    wstr = wstr & " -M.融資金額+M.元金額+M.利息額 As 合計額"
    
    wstr = wstr & " FROM ((DCKA010_資金繰表 AS M"
    wstr = wstr & " INNER JOIN DBDA010_借入金 AS K"
    wstr = wstr & " ON M.借入番号 = K.借入番号)"
    wstr = wstr & " INNER JOIN DAAA040_銀行マスタ AS G"
    wstr = wstr & " ON K.銀行番号 = G.銀行番号)"
    wstr = wstr & " INNER JOIN DAAA116_借入金種別 AS S"
    wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分"
    
    GVar1 = C年月日.平成To西暦("年月日", GRpt.テキスト_01)
    GVar2 = C年月日.平成To西暦("年月日", GRpt.テキスト_02)
    
    wstr = wstr & " WHERE ((M.返済金額<>0"
    wstr = wstr & " AND M.融資残高>=0)"
    wstr = wstr & " OR M.融資金額<>0)"
    
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
    
    If GRpt.集計 = "借入金種別区分集計しない" Then
        wstr = wstr & " ORDER BY K.銀行番号,K.長短区分 desc,K.金利種別 desc,M.実際年月日,M.借入番号"
    Else
        wstr = wstr & " ORDER BY K.銀行番号,K.借入金種別区分,K.長短区分 desc,K.金利種別 desc,M.実際年月日,M.借入番号"
    End If
   
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
            
        '名称
        Write #1, _
            "銀行名", "長短区分", "金利種別", "借入番号", "支払日", "融資金額", "元金額", "利息額", "合計額"
            
    Do Until wRs.eof
        
        Write #1, _
            P8.FCStr(wRs.Fields("銀行名").Value), _
            P8.FCStr(wRs.Fields("長短区分").Value), _
            P8.FCStr(wRs.Fields("金利種別").Value), _
            P8.FCStr(wRs.Fields("借入番号").Value), _
            P8.FCStr(wRs.Fields("支払日").Value), _
            P8.FCDbl(wRs.Fields("融資金額").Value), _
            P8.FCDbl(wRs.Fields("元金額").Value), _
            P8.FCDbl(wRs.Fields("利息額").Value), _
            P8.FCDbl(wRs.Fields("合計額").Value)
            
        wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_1年内返済集計表_杉村倉庫
'------------------------------------------------
Private Sub MX040_1年内返済集計表_杉村倉庫(pCsvFileName As String)
'
    Dim ws01 As String
    Dim wstr2 As String
    Dim w開始年月日 As Date
    Dim w推移表区分 As String
    Dim w番号 As String, w番号2 As String, w番号3 As String
    Dim ws銀行番号 As String, ws銀行名 As String
    Dim WKGCNT As Integer
    Dim wc_融資金額 As Currency, wc_当期融資残高 As Currency, wc_来前期融資残高 As Currency, wc_来期融資残高 As Currency
    Dim wc_1年内返済額 As Currency, wc_追加計上額 As Currency
    Dim wOrder As String
    Dim FLG_Order As Boolean
    Dim wdate As Date
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "西暦") '年度変換後
    w推移表タイトル = MUA010_推移表ファイル作成("", "", w開始年月日, GRpt.推移, 12)
    
    w番号 = Right("00" + CStr(GInt1), 2)
    w番号2 = Right("00" + CStr(GInt1 + 4), 2)
    w番号3 = Right("00" + CStr(GInt1 + 3), 2)

    wdate = DateAdd("yyyy", 1, w開始年月日)
'
    '** 合計表示 **
    If GRpt.詳細表示 = 0 Then
        wstr = "SELECT "
        wstr = wstr & "count(K.銀行番号) As 件数,"
        wstr = wstr & "K.銀行番号,"
        wstr = wstr & "MIN(G.銀行名) As 銀行名,"
        wstr = wstr & "SUM(K.融資金額) As 融資金額,"
        wstr = wstr & "SUM(残高_" + w番号 + ") As 当期融資残高,"
        wstr = wstr & "SUM(残高_" + w番号3 + ") As 来前期融資残高,"
        wstr = wstr & "SUM(残高_" + w番号2 + ") As 来期融資残高,"
        wstr = wstr & "SUM(長短振替額_" + w番号 + ") As 1年内返済額,"
        wstr = wstr & "SUM(残高_" + w番号3 + "-残高_" + w番号2 + ") As 追加計上額"
    
        wstr = wstr & " FROM ((((((DCDA010_借入残高推移表結果 As Z"
        wstr = wstr & " INNER JOIN DCDA010_借入残高推移表結果2 As Z2"
        wstr = wstr & " ON Z.借入番号=Z2.借入番号)"
        wstr = wstr & " INNER JOIN DCIA010_借入金ワーク As K"
        wstr = wstr & " ON Z.借入番号=K.借入番号)"
        wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
        wstr = wstr & " ON K.銀行番号=G.銀行番号)"
        wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
        wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分)"
        wstr = wstr & " LEFT JOIN DAAA200_部門マスタ As B"
        wstr = wstr & " ON K.プロジェクト番号 = B.部門番号)"
        wstr = wstr + " Left Join DAAA116_基準金利 As KK"
        wstr = wstr + " ON K.基準金利区分 = KK.基準金利区分)"
        
        'Where 条件
        'All 0 は表示しない
        ws01 = Right("00" + CStr(GInt1), 2)
        wstr = wstr & " Where ("
        wstr = wstr & " Z.融資_" & ws01 & "<>0"
        wstr = wstr & " Or Z.元金_" & ws01 & "<>0"
        wstr = wstr & " Or Z.利息_" & ws01 & "<>0"
        wstr = wstr & " Or Z.返済_" & ws01 & "<>0"
        wstr = wstr & " Or Z.解約_" & ws01 & "<>0"
        wstr = wstr & " Or Z.保証_" & ws01 & "<>0"
        wstr = wstr & " Or Z.残高_" & ws01 & "<>0"
        wstr = wstr & " Or Z.前払利息_" & ws01 & "<>0"
        wstr = wstr & " Or Z.前払利息増_" & ws01 & "<>0"
        wstr = wstr & " Or Z.前払利息減_" & ws01 & "<>0"
        wstr = wstr & " Or Z.未払利息_" & ws01 & "<>0"
        wstr = wstr & " Or Z.未払利息増_" & ws01 & "<>0"
        wstr = wstr & " Or Z.未払利息減_" & ws01 & "<>0"
        wstr = wstr & " Or Z2.損益利息額_" & ws01 & "<>0"
        wstr = wstr & " )"
        
        wstr = wstr & " Group by K.銀行番号"
        wstr = wstr & " Order by K.銀行番号"
        
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        If Not wRs.eof Then
        
        Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
        
        On Error GoTo Err_Hundle
            
            Write #1, _
                "銀行番号", _
                "銀行名", _
                "件数", _
                "融資金額", _
                GRpt.テキスト_01 & "年" & CInt(Right(w推移表タイトル.X番目年月(GInt1), 2)) & "月" & "融資残高", _
                Format(wdate, "yyyy") & "年" & CInt(Right(w推移表タイトル.X番目年月(GInt1 + 3), 2)) & "月" & "融資残高", _
                Format(wdate, "yyyy") & "年" & CInt(Right(w推移表タイトル.X番目年月(GInt1 + 4), 2)) & "月" & "融資残高", _
                "1年内返済額", _
                "追加計上額"
                
            Do Until wRs.eof
                Write #1, _
                    P8.FCStr(wRs.Fields("銀行番号").Value), _
                    P8.FCStr(wRs.Fields("銀行名").Value), _
                    P8.FCDbl(wRs.Fields("件数").Value), _
                    P8.FCDbl(wRs.Fields("融資金額").Value), _
                    P8.FCDbl(wRs.Fields("当期融資残高").Value), _
                    P8.FCDbl(wRs.Fields("来前期融資残高").Value), _
                    P8.FCDbl(wRs.Fields("来期融資残高").Value), _
                    P8.FCDbl(wRs.Fields("1年内返済額").Value), _
                    P8.FCDbl(wRs.Fields("追加計上額").Value)
               
                wRs.MoveNext
            Loop
        End If
        wRs.Close
        Set wRs = Nothing
        
        '合計
        wstr2 = "select "
        wstr2 = wstr2 & "count(Z.借入番号) As 件数,"
        wstr2 = wstr2 & "Sum(K.融資金額) As 融資金額,"
        wstr2 = wstr2 & "Sum(残高_" + w番号 + ") As 当期融資残高,"
        wstr2 = wstr2 & "Sum(残高_" + w番号3 + ") As 来前期融資残高,"
        wstr2 = wstr2 & "Sum(残高_" + w番号2 + ") As 来期融資残高,"
        wstr2 = wstr2 & "Sum(長短振替額_" + w番号 + ") As 1年内返済額,"
        wstr2 = wstr2 & "Sum(残高_" + w番号3 + ")-Sum(残高_" + w番号2 + ") As 追加計上額"
        wstr2 = wstr2 & " FROM ((((((DCDA010_借入残高推移表結果 As Z"
        wstr2 = wstr2 & " INNER JOIN DCDA010_借入残高推移表結果2 As Z2"
        wstr2 = wstr2 & " ON Z.借入番号=Z2.借入番号)"
        wstr2 = wstr2 & " INNER JOIN DCIA010_借入金ワーク As K"
        wstr2 = wstr2 & " ON Z.借入番号=K.借入番号)"
        wstr2 = wstr2 & " INNER JOIN DAAA040_銀行マスタ As G"
        wstr2 = wstr2 & " ON K.銀行番号=G.銀行番号)"
        wstr2 = wstr2 & " LEFT JOIN DAAA116_借入金種別 As S"
        wstr2 = wstr2 & " ON K.借入金種別区分 = S.借入金種別区分)"
        wstr2 = wstr2 & " LEFT JOIN DAAA200_部門マスタ As B"
        wstr2 = wstr2 & " ON K.プロジェクト番号 = B.部門番号)"
        wstr2 = wstr2 + " Left Join DAAA116_基準金利 As KK"
        wstr2 = wstr2 + " ON K.基準金利区分 = KK.基準金利区分)"
        
        'Where 条件
        'All 0 は表示しない
        ws01 = Right("00" + CStr(GInt1), 2)
        wstr2 = wstr2 & " Where ("
        wstr2 = wstr2 & " Z.融資_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.元金_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.利息_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.返済_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.解約_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.保証_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.残高_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.前払利息_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.前払利息増_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.前払利息減_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.未払利息_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.未払利息増_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.未払利息減_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z2.損益利息額_" & ws01 & "<>0"
        wstr2 = wstr2 & " )"
    
        Call AdoRecordsetOpen(GDb, wRs, wstr2)
        If Not wRs.eof Then
        Do Until wRs.eof
        
            Write #1, _
                "総合計", "", _
                P8.FCDbl(wRs.Fields("件数").Value), _
                P8.FCDbl(wRs.Fields("融資金額").Value), _
                P8.FCDbl(wRs.Fields("当期融資残高").Value), _
                P8.FCDbl(wRs.Fields("来前期融資残高").Value), _
                P8.FCDbl(wRs.Fields("来期融資残高").Value), _
                P8.FCDbl(wRs.Fields("1年内返済額").Value), _
                P8.FCDbl(wRs.Fields("追加計上額").Value)
    
        wRs.MoveNext
        Loop
        End If
        wRs.Close
        Set wRs = Nothing
            
        Close #1 '出力ファイルを閉じる
    
    '** 詳細表示 **
    Else
        wstr = "Select "
        wstr = wstr & "K.借入番号,"
        
        'セクションGR
        wstr = wstr & "K.銀行番号,"
        wstr = wstr & "G.金融機関番号,"
        wstr = wstr & "B.部門番号,"
        wstr = wstr & "K.借入金種別区分,"
        
        wstr = wstr & "G.銀行名,"
        wstr = wstr & "G.金融機関名,"
        wstr = wstr & "B.部門名,"
        wstr = wstr & "S.借入金種別名,"
        
        wstr = wstr & "K.融資金額,"
        wstr = wstr & "利率_" + w番号 + " As 利率,"
        wstr = wstr & "残高_" + w番号 + " As 当期融資残高,"
        wstr = wstr & "残高_" + w番号3 + " As 来前期融資残高,"
        wstr = wstr & "残高_" + w番号2 + " As 来期融資残高,"
        wstr = wstr & "長短振替額_" + w番号 + " As 1年内返済額,"
        wstr = wstr & "残高_" + w番号3 + "-残高_" + w番号2 + " As 追加計上額"
    
        wstr = wstr & " FROM ((((((DCDA010_借入残高推移表結果 As Z"
        wstr = wstr & " INNER JOIN DCDA010_借入残高推移表結果2 As Z2"
        wstr = wstr & " ON Z.借入番号=Z2.借入番号)"
        wstr = wstr & " INNER JOIN DCIA010_借入金ワーク As K"
        wstr = wstr & " ON Z.借入番号=K.借入番号)"
        wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
        wstr = wstr & " ON K.銀行番号=G.銀行番号)"
        wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
        wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分)"
        wstr = wstr & " LEFT JOIN DAAA200_部門マスタ As B"
        wstr = wstr & " ON K.プロジェクト番号 = B.部門番号)"
        wstr = wstr + " Left Join DAAA116_基準金利 As KK"
        wstr = wstr + " ON K.基準金利区分 = KK.基準金利区分)"
        
        'Where 条件
        'All 0 は表示しない
        ws01 = Right("00" + CStr(GInt1), 2)
        wstr = wstr & " Where ("
        wstr = wstr & " Z.融資_" & ws01 & "<>0"
        wstr = wstr & " Or Z.元金_" & ws01 & "<>0"
        wstr = wstr & " Or Z.利息_" & ws01 & "<>0"
        wstr = wstr & " Or Z.返済_" & ws01 & "<>0"
        wstr = wstr & " Or Z.解約_" & ws01 & "<>0"
        wstr = wstr & " Or Z.保証_" & ws01 & "<>0"
        wstr = wstr & " Or Z.残高_" & ws01 & "<>0"
        wstr = wstr & " Or Z.前払利息_" & ws01 & "<>0"
        wstr = wstr & " Or Z.前払利息増_" & ws01 & "<>0"
        wstr = wstr & " Or Z.前払利息減_" & ws01 & "<>0"
        wstr = wstr & " Or Z.未払利息_" & ws01 & "<>0"
        wstr = wstr & " Or Z.未払利息増_" & ws01 & "<>0"
        wstr = wstr & " Or Z.未払利息減_" & ws01 & "<>0"
        wstr = wstr & " Or Z2.損益利息額_" & ws01 & "<>0"
        wstr = wstr & " )"
        
        wstr = wstr & " Order by K.銀行番号,K.借入番号"
        
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        If Not wRs.eof Then
        
        Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
        
        On Error GoTo Err_Hundle
            
            Write #1, _
                "借入番号", _
                "銀行番号", _
                "銀行名", _
                "利率/件数", _
                "融資金額", _
                GRpt.テキスト_01 & "年" & CInt(Right(w推移表タイトル.X番目年月(GInt1), 2)) & "月" & "融資残高", _
                Format(wdate, "yyyy") & "年" & CInt(Right(w推移表タイトル.X番目年月(GInt1 + 3), 2)) & "月" & "融資残高", _
                Format(wdate, "yyyy") & "年" & CInt(Right(w推移表タイトル.X番目年月(GInt1 + 4), 2)) & "月" & "融資残高", _
                "1年内返済額", _
                "追加計上額"
                
            ws銀行番号 = "": ws銀行名 = ""
            
            Do Until wRs.eof
                
                '銀行計
                If ws銀行番号 <> "" And ws銀行番号 <> P8.FCStr(wRs.Fields("銀行番号").Value) Then
                    Write #1, _
                        "小計", _
                        ws銀行番号, _
                        ws銀行名, _
                        WKGCNT, _
                        wc_融資金額, _
                        wc_当期融資残高, _
                        wc_来前期融資残高, _
                        wc_来期融資残高, _
                        wc_1年内返済額, _
                        wc_追加計上額
                                
                        WKGCNT = 0
                        wc_融資金額 = 0
                        wc_当期融資残高 = 0
                        wc_来前期融資残高 = 0
                        wc_来期融資残高 = 0
                        wc_1年内返済額 = 0
                        wc_追加計上額 = 0
                End If
                
                Write #1, _
                    P8.FCStr(wRs.Fields("借入番号").Value), _
                    P8.FCStr(wRs.Fields("銀行番号").Value), _
                    P8.FCStr(wRs.Fields("銀行名").Value), _
                    P8.FCDbl(wRs.Fields("利率").Value), _
                    P8.FCDbl(wRs.Fields("融資金額").Value), _
                    P8.FCDbl(wRs.Fields("当期融資残高").Value), _
                    P8.FCDbl(wRs.Fields("来前期融資残高").Value), _
                    P8.FCDbl(wRs.Fields("来期融資残高").Value), _
                    P8.FCDbl(wRs.Fields("1年内返済額").Value), _
                    P8.FCDbl(wRs.Fields("追加計上額").Value)
               
                ws銀行番号 = P8.FCStr(wRs.Fields("銀行番号").Value)
                ws銀行名 = P8.FCStr(wRs.Fields("銀行名").Value)
                
                '銀行計
                WKGCNT = WKGCNT + 1
                wc_融資金額 = wc_融資金額 + P8.FCDbl(wRs.Fields("融資金額").Value)
                wc_当期融資残高 = wc_当期融資残高 + P8.FCDbl(wRs.Fields("当期融資残高").Value)
                wc_来前期融資残高 = wc_来前期融資残高 + P8.FCDbl(wRs.Fields("来前期融資残高").Value)
                wc_来期融資残高 = wc_来期融資残高 + P8.FCDbl(wRs.Fields("来期融資残高").Value)
                wc_1年内返済額 = wc_1年内返済額 + P8.FCDbl(wRs.Fields("1年内返済額").Value)
                wc_追加計上額 = wc_追加計上額 + P8.FCDbl(wRs.Fields("追加計上額").Value)

                wRs.MoveNext
            Loop
            
            '銀行計
            If ws銀行番号 <> "" Then
                Write #1, _
                    "小計", _
                    ws銀行番号, _
                    ws銀行名, _
                    WKGCNT, _
                    wc_融資金額, _
                    wc_当期融資残高, _
                    wc_来前期融資残高, _
                    wc_来期融資残高, _
                    wc_1年内返済額, _
                    wc_追加計上額
            End If
    
        End If
        wRs.Close
        Set wRs = Nothing
        
        '合計
        wstr2 = "select "
        wstr2 = wstr2 & "count(Z.借入番号) As 件数,"
        wstr2 = wstr2 & "Sum(K.融資金額) As 融資金額,"
        wstr2 = wstr2 & "Sum(残高_" + w番号 + ") As 当期融資残高,"
        wstr2 = wstr2 & "Sum(残高_" + w番号3 + ") As 来前期融資残高,"
        wstr2 = wstr2 & "Sum(残高_" + w番号2 + ") As 来期融資残高,"
        wstr2 = wstr2 & "Sum(長短振替額_" + w番号 + ") As 1年内返済額,"
        wstr2 = wstr2 & "Sum(残高_" + w番号3 + ")-Sum(残高_" + w番号2 + ") As 追加計上額"
        wstr2 = wstr2 & " FROM ((((((DCDA010_借入残高推移表結果 As Z"
        wstr2 = wstr2 & " INNER JOIN DCDA010_借入残高推移表結果2 As Z2"
        wstr2 = wstr2 & " ON Z.借入番号=Z2.借入番号)"
        wstr2 = wstr2 & " INNER JOIN DCIA010_借入金ワーク As K"
        wstr2 = wstr2 & " ON Z.借入番号=K.借入番号)"
        wstr2 = wstr2 & " INNER JOIN DAAA040_銀行マスタ As G"
        wstr2 = wstr2 & " ON K.銀行番号=G.銀行番号)"
        wstr2 = wstr2 & " LEFT JOIN DAAA116_借入金種別 As S"
        wstr2 = wstr2 & " ON K.借入金種別区分 = S.借入金種別区分)"
        wstr2 = wstr2 & " LEFT JOIN DAAA200_部門マスタ As B"
        wstr2 = wstr2 & " ON K.プロジェクト番号 = B.部門番号)"
        wstr2 = wstr2 + " Left Join DAAA116_基準金利 As KK"
        wstr2 = wstr2 + " ON K.基準金利区分 = KK.基準金利区分)"
        
        'Where 条件
        'All 0 は表示しない
        ws01 = Right("00" + CStr(GInt1), 2)
        wstr2 = wstr2 & " Where ("
        wstr2 = wstr2 & " Z.融資_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.元金_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.利息_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.返済_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.解約_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.保証_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.残高_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.前払利息_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.前払利息増_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.前払利息減_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.未払利息_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.未払利息増_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.未払利息減_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z2.損益利息額_" & ws01 & "<>0"
        wstr2 = wstr2 & " )"
    
        Call AdoRecordsetOpen(GDb, wRs, wstr2)
        If Not wRs.eof Then
        Do Until wRs.eof
        
            Write #1, _
                "総合計", "", "", _
                P8.FCDbl(wRs.Fields("件数").Value), _
                P8.FCDbl(wRs.Fields("融資金額").Value), _
                P8.FCDbl(wRs.Fields("当期融資残高").Value), _
                P8.FCDbl(wRs.Fields("来前期融資残高").Value), _
                P8.FCDbl(wRs.Fields("来期融資残高").Value), _
                P8.FCDbl(wRs.Fields("1年内返済額").Value), _
                P8.FCDbl(wRs.Fields("追加計上額").Value)
    
        wRs.MoveNext
        Loop
        End If
        wRs.Close
        Set wRs = Nothing
            
        Close #1 '出力ファイルを閉じる
    End If
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_銀行別利息表_杉村倉庫
'------------------------------------------------
Private Sub MX040_銀行別利息表_杉村倉庫(pCsvFileName As String, pRisokuKBN As String)
'
    Dim ws01 As String, wstr2 As String
    Dim w開始年月日 As Date
    Dim w推移表区分 As String
    Dim w番号 As String
    Dim wOrder As String
    Dim ws銀行番号 As String, ws銀行名 As String
    Dim WKGCNT As Integer
    Dim wc_融資金額 As Currency, wc_融資残高 As Currency, wc_支払額 As Currency, wc_洗替 As Currency, wc_計上 As Currency, wc_支払利息 As Currency
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "西暦") '年度変換後
    w推移表タイトル = MUA010_推移表ファイル作成("", "", w開始年月日, GRpt.推移, 12)

    w番号 = Right("00" + CStr(GInt1), 2)
'
    '** 合計表示 **
    If GRpt.詳細表示 = 0 Then
        wstr = "Select "
        wstr = wstr & "K.銀行番号,"
        wstr = wstr & "MIN(G.銀行名) As 銀行名,"
        wstr = wstr & "count(Z.借入番号) As 件数,"
        wstr = wstr & "SUM(K.融資金額) As 融資金額,"
        wstr = wstr & "SUM(Z.残高_" & w番号 & ") As 融資残高,"
        '杉村倉庫仕様
        wstr = wstr & "IIF(MIN(K.利息区分)='" & XMXA020_区分("利息区分", "利息先払") & "',SUM(Z.前払利息増_" & w番号 & "),SUM(Z.未払利息減_" & w番号 & ")) As 支払額,"
        wstr = wstr & "IIF(MIN(K.利息区分)='" & XMXA020_区分("利息区分", "利息先払") & "',SUM(Z.前払利息減_" & w番号 & "-Z.前払利息増_" & w番号 & "+Z.前払利息_" & w番号 & "),SUM(-Z.未払利息増_" & w番号 & "+Z.未払利息減_" & w番号 & "+Z.未払利息_" & w番号 & ")) As 洗替,"
        wstr = wstr & "IIF(MIN(K.利息区分)='" & XMXA020_区分("利息区分", "利息先払") & "',SUM(Z.前払利息_" & w番号 & "),SUM(Z.未払利息_" & w番号 & ")) As 計上,"
        wstr = wstr & "IIF(MIN(K.利息区分)='" & XMXA020_区分("利息区分", "利息先払") & "',SUM(Z.前払利息減_" & w番号 & "),SUM(Z.未払利息増_" & w番号 & ")) As 支払利息"
        
        wstr = wstr & " FROM ((((DCDA010_借入残高推移表結果 As Z"
        wstr = wstr & " INNER JOIN DCDA010_借入残高推移表結果２ As Z2"
        wstr = wstr & " ON Z.借入番号=Z2.借入番号)"
        wstr = wstr & " INNER JOIN DCIA010_借入金ワーク As K"
        wstr = wstr & " ON Z.借入番号=K.借入番号)"
        wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
        wstr = wstr & " ON K.銀行番号=G.銀行番号)"
        wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
        wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分)"
        
        'Where 条件
        If pRisokuKBN = "利息先払" Then
            wstr = wstr & " Where K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "'"
        Else
            wstr = wstr & " Where K.利息区分='" & XMXA020_区分("利息区分", "利息後払") & "'"
        End If
        
        'All 0 は表示しない
        ws01 = Right("00" + CStr(GInt1), 2)
        wstr = wstr & " AND ("
        wstr = wstr & " Z.融資_" & ws01 & "<>0"
        wstr = wstr & " Or Z.元金_" & ws01 & "<>0"
        wstr = wstr & " Or Z.利息_" & ws01 & "<>0"
        wstr = wstr & " Or Z.返済_" & ws01 & "<>0"
        wstr = wstr & " Or Z.解約_" & ws01 & "<>0"
        wstr = wstr & " Or Z.保証_" & ws01 & "<>0"
        wstr = wstr & " Or Z.残高_" & ws01 & "<>0"
        wstr = wstr & " Or Z.前払利息_" & ws01 & "<>0"
        wstr = wstr & " Or Z.前払利息増_" & ws01 & "<>0"
        wstr = wstr & " Or Z.前払利息減_" & ws01 & "<>0"
        wstr = wstr & " Or Z.未払利息_" & ws01 & "<>0"
        wstr = wstr & " Or Z.未払利息増_" & ws01 & "<>0"
        wstr = wstr & " Or Z.未払利息減_" & ws01 & "<>0"
        wstr = wstr & " Or Z2.損益利息額_" & ws01 & "<>0"
        wstr = wstr & " )"
        
        wstr = wstr & " Group by K.銀行番号"
        wstr = wstr & " Order by K.銀行番号"
    
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        If Not wRs.eof Then
        
        Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
        
        On Error GoTo Err_Hundle
            
            '名称
            If pRisokuKBN = "利息先払" Then
                Write #1, _
                    "銀行番号", "銀行名", "件数", "融資金額", "融資残高", "支払額", "前払利息(洗替+)", "前払利息(洗替+)", "支払利息"
            Else
                Write #1, _
                    "銀行番号", "銀行名", "件数", "融資金額", "融資残高", "支払額", "未払利息(洗替-)", "未払利息(計上額+)", "支払利息"
            End If
        
            Do Until wRs.eof
            
                '銀行計
                Write #1, _
                    P8.FCStr(wRs.Fields("銀行番号").Value), _
                    P8.FCStr(wRs.Fields("銀行名").Value), _
                    wRs.Fields("件数").Value, _
                    P8.FCDbl(wRs.Fields("融資金額").Value), _
                    P8.FCDbl(wRs.Fields("融資残高").Value), _
                    P8.FCDbl(wRs.Fields("支払額").Value), _
                    P8.FCDbl(wRs.Fields("洗替").Value), _
                    P8.FCDbl(wRs.Fields("計上").Value), _
                    P8.FCDbl(wRs.Fields("支払利息").Value)
            
            wRs.MoveNext
            Loop
            
        End If
        wRs.Close
        Set wRs = Nothing
        
        '合計
        wstr2 = "select "
        wstr2 = wstr2 & "count(Z.借入番号) As 件数,"
        wstr2 = wstr2 & "SUM(K.融資金額) As 融資金額,"
        wstr2 = wstr2 & "SUM(Z.残高_" & w番号 & ") As 融資残高,"
        wstr2 = wstr2 & "SUM(IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息増_" & w番号 & ",Z.未払利息減_" & w番号 & ")) As 支払額,"
        wstr2 = wstr2 & "SUM(IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & "-Z.前払利息増_" & w番号 & "+Z.前払利息_" & w番号 & ",-Z.未払利息増_" & w番号 & "+Z.未払利息減_" & w番号 & "+Z.未払利息_" & w番号 & ")) As 洗替,"
        wstr2 = wstr2 & "SUM(IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息_" & w番号 & ",Z.未払利息_" & w番号 & ")) As 計上,"
        wstr2 = wstr2 & "SUM(IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & ",Z.未払利息増_" & w番号 & ")) As 支払利息"
        wstr2 = wstr2 + " FROM (DCDA010_借入残高推移表結果 As Z"
        wstr2 = wstr2 + " Inner Join DCDA010_借入残高推移表結果2 As Z2"
        wstr2 = wstr2 + " ON Z.借入番号 = Z2.借入番号)"
        wstr2 = wstr2 + " Inner Join DCIA010_借入金ワーク As K"
        wstr2 = wstr2 + " ON Z.借入番号 = K.借入番号"
        
        'Where 条件
        If pRisokuKBN = "利息先払" Then
            wstr2 = wstr2 & " Where K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "'"
        Else
            wstr2 = wstr2 & " Where K.利息区分='" & XMXA020_区分("利息区分", "利息後払") & "'"
        End If
    
        'All 0 は表示しない
        ws01 = Right("00" + CStr(GInt1), 2)
        wstr2 = wstr2 & " AND ("
        wstr2 = wstr2 & " Z.融資_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.元金_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.利息_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.返済_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.解約_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.保証_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.残高_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.前払利息_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.前払利息増_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.前払利息減_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.未払利息_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.未払利息増_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.未払利息減_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z2.損益利息額_" & ws01 & "<>0"
        wstr2 = wstr2 & " )"
    
        Call AdoRecordsetOpen(GDb, wRs, wstr2)
        If Not wRs.eof Then
        Do Until wRs.eof
        
            If pRisokuKBN = "利息先払" Then
                ws01 = "利息先払計"
            Else
                ws01 = "利息後払計"
            End If
                
            Write #1, _
                ws01, "", wRs.Fields("件数").Value, _
                P8.FCDbl(wRs.Fields("融資金額").Value), _
                P8.FCDbl(wRs.Fields("融資残高").Value), _
                P8.FCDbl(wRs.Fields("支払額").Value), _
                P8.FCDbl(wRs.Fields("洗替").Value), _
                P8.FCDbl(wRs.Fields("計上").Value), _
                P8.FCDbl(wRs.Fields("支払利息").Value)
    
        wRs.MoveNext
        Loop
        End If
        wRs.Close
        Set wRs = Nothing
            
        Close #1 '出力ファイルを閉じる
        
    '** 詳細表示 **
    Else
        wstr = "Select "
        wstr = wstr & "K.借入番号 As 借入番号,"
        wstr = wstr & "K.銀行番号,"
        wstr = wstr & "G.銀行名 As 銀行名,"
        wstr = wstr & "Z.利率_" & w番号 & " As 利率,"
        wstr = wstr & "K.融資金額 As 融資金額,"
        wstr = wstr & "Z.残高_" & w番号 & " As 融資残高,"
        '杉村倉庫仕様
        wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息増_" & w番号 & ",Z.未払利息減_" & w番号 & ") As 支払額,"
        wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & "-Z.前払利息増_" & w番号 & "+Z.前払利息_" & w番号 & ",-Z.未払利息増_" & w番号 & "+Z.未払利息減_" & w番号 & "+Z.未払利息_" & w番号 & ") As 洗替,"
        wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息_" & w番号 & ",Z.未払利息_" & w番号 & ") As 計上,"
        wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & ",Z.未払利息増_" & w番号 & ") As 支払利息"
        
        wstr = wstr & " FROM ((((DCDA010_借入残高推移表結果 As Z"
        wstr = wstr & " INNER JOIN DCDA010_借入残高推移表結果２ As Z2"
        wstr = wstr & " ON Z.借入番号=Z2.借入番号)"
        wstr = wstr & " INNER JOIN DCIA010_借入金ワーク As K"
        wstr = wstr & " ON Z.借入番号=K.借入番号)"
        wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
        wstr = wstr & " ON K.銀行番号=G.銀行番号)"
        wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
        wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分)"
        
        'Where 条件
        If pRisokuKBN = "利息先払" Then
            wstr = wstr & " Where K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "'"
        Else
            wstr = wstr & " Where K.利息区分='" & XMXA020_区分("利息区分", "利息後払") & "'"
        End If
                
        'All 0 は表示しない
        ws01 = Right("00" + CStr(GInt1), 2)
        wstr = wstr & " AND ("
        wstr = wstr & " Z.融資_" & ws01 & "<>0"
        wstr = wstr & " Or Z.元金_" & ws01 & "<>0"
        wstr = wstr & " Or Z.利息_" & ws01 & "<>0"
        wstr = wstr & " Or Z.返済_" & ws01 & "<>0"
        wstr = wstr & " Or Z.解約_" & ws01 & "<>0"
        wstr = wstr & " Or Z.保証_" & ws01 & "<>0"
        wstr = wstr & " Or Z.残高_" & ws01 & "<>0"
        wstr = wstr & " Or Z.前払利息_" & ws01 & "<>0"
        wstr = wstr & " Or Z.前払利息増_" & ws01 & "<>0"
        wstr = wstr & " Or Z.前払利息減_" & ws01 & "<>0"
        wstr = wstr & " Or Z.未払利息_" & ws01 & "<>0"
        wstr = wstr & " Or Z.未払利息増_" & ws01 & "<>0"
        wstr = wstr & " Or Z.未払利息減_" & ws01 & "<>0"
        wstr = wstr & " Or Z2.損益利息額_" & ws01 & "<>0"
        wstr = wstr & " )"
                
        wOrder = " Order by K.銀行番号,K.借入番号"
        wstr = wstr & wOrder
        Call AdoRecordsetOpen(GDb, wRs, wstr)
        If Not wRs.eof Then
        
        Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
        
        On Error GoTo Err_Hundle
            
            '名称
            If pRisokuKBN = "利息先払" Then
                Write #1, _
                    "借入番号", "銀行番号", "銀行名", "利率/件数", "融資金額", "融資残高", "支払額", "前払利息(洗替+)", "前払利息(洗替+)", "支払利息"
            Else
                Write #1, _
                    "借入番号", "銀行番号", "銀行名", "利率/件数", "融資金額", "融資残高", "支払額", "未払利息(洗替-)", "未払利息(計上額+)", "支払利息"
            End If
            
            ws銀行番号 = "": ws銀行名 = ""
            
            Do Until wRs.eof
            
                '銀行計
                If ws銀行番号 <> "" And ws銀行番号 <> P8.FCStr(wRs.Fields("銀行番号").Value) Then
                    Write #1, _
                        "小計", _
                        ws銀行番号, _
                        ws銀行名, _
                        WKGCNT, _
                        wc_融資金額, _
                        wc_融資残高, _
                        wc_支払額, _
                        wc_洗替, _
                        wc_計上, _
                        wc_支払利息
                                
                        WKGCNT = 0
                        wc_融資金額 = 0
                        wc_融資残高 = 0
                        wc_支払額 = 0
                        wc_洗替 = 0
                        wc_計上 = 0
                        wc_支払利息 = 0
                End If
                               
                Write #1, _
                    P8.FCStr(wRs.Fields("借入番号").Value), _
                    P8.FCStr(wRs.Fields("銀行番号").Value), _
                    P8.FCStr(wRs.Fields("銀行名").Value), _
                    P8.FCDbl(wRs.Fields("利率").Value), _
                    P8.FCStr(wRs.Fields("融資金額").Value), _
                    P8.FCDbl(wRs.Fields("融資残高").Value), _
                    P8.FCDbl(wRs.Fields("支払額").Value), _
                    P8.FCDbl(wRs.Fields("洗替").Value), _
                    P8.FCDbl(wRs.Fields("計上").Value), _
                    P8.FCDbl(wRs.Fields("支払利息").Value)
                
                ws銀行番号 = P8.FCStr(wRs.Fields("銀行番号").Value)
                ws銀行名 = P8.FCStr(wRs.Fields("銀行名").Value)
                
                '銀行計
                WKGCNT = WKGCNT + 1
                wc_融資金額 = wc_融資金額 + wRs.Fields("融資金額").Value
                wc_融資残高 = wc_融資残高 + wRs.Fields("融資残高").Value
                wc_支払額 = wc_支払額 + wRs.Fields("支払額").Value
                wc_洗替 = wc_洗替 + wRs.Fields("洗替").Value
                wc_計上 = wc_計上 + wRs.Fields("計上").Value
                wc_支払利息 = wc_支払利息 + wRs.Fields("支払利息").Value
                
            wRs.MoveNext
            Loop
            
           '銀行計
            If ws銀行番号 <> "" Then
                Write #1, _
                    "小計", _
                    ws銀行番号, _
                    ws銀行名, _
                    WKGCNT, _
                    wc_融資金額, _
                    wc_融資残高, _
                    wc_支払額, _
                    wc_洗替, _
                    wc_計上, _
                    wc_支払利息
            End If
                
        End If
        wRs.Close
        Set wRs = Nothing
        
        '合計
        wstr2 = "select "
        wstr2 = wstr2 & "count(Z.借入番号) As 件数,"
        wstr2 = wstr2 & "SUM(K.融資金額) As 融資金額,"
        wstr2 = wstr2 & "SUM(Z.残高_" & w番号 & ") As 融資残高,"
        wstr2 = wstr2 & "SUM(IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息増_" & w番号 & ",Z.未払利息減_" & w番号 & ")) As 支払額,"
        wstr2 = wstr2 & "SUM(IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & "-Z.前払利息増_" & w番号 & "+Z.前払利息_" & w番号 & ",-Z.未払利息増_" & w番号 & "+Z.未払利息減_" & w番号 & "+Z.未払利息_" & w番号 & ")) As 洗替,"
        wstr2 = wstr2 & "SUM(IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息_" & w番号 & ",Z.未払利息_" & w番号 & ")) As 計上,"
        wstr2 = wstr2 & "SUM(IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & ",Z.未払利息増_" & w番号 & ")) As 支払利息"
        wstr2 = wstr2 + " FROM (DCDA010_借入残高推移表結果 As Z"
        wstr2 = wstr2 + " Inner Join DCDA010_借入残高推移表結果2 As Z2"
        wstr2 = wstr2 + " ON Z.借入番号 = Z2.借入番号)"
        wstr2 = wstr2 + " Inner Join DCIA010_借入金ワーク As K"
        wstr2 = wstr2 + " ON Z.借入番号 = K.借入番号"
        
        'Where 条件
        If pRisokuKBN = "利息先払" Then
            wstr2 = wstr2 & " Where K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "'"
        Else
            wstr2 = wstr2 & " Where K.利息区分='" & XMXA020_区分("利息区分", "利息後払") & "'"
        End If
    
        'All 0 は表示しない
        ws01 = Right("00" + CStr(GInt1), 2)
        wstr2 = wstr2 & " AND ("
        wstr2 = wstr2 & " Z.融資_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.元金_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.利息_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.返済_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.解約_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.保証_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.残高_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.前払利息_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.前払利息増_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.前払利息減_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.未払利息_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.未払利息増_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z.未払利息減_" & ws01 & "<>0"
        wstr2 = wstr2 & " Or Z2.損益利息額_" & ws01 & "<>0"
        wstr2 = wstr2 & " )"
    
        Call AdoRecordsetOpen(GDb, wRs, wstr2)
        If Not wRs.eof Then
        Do Until wRs.eof
        
            If pRisokuKBN = "利息先払" Then
                ws01 = "利息先払計"
            Else
                ws01 = "利息後払計"
            End If
            
            Write #1, _
                ws01, "", "", wRs.Fields("件数").Value, _
                P8.FCDbl(wRs.Fields("融資金額").Value), _
                P8.FCDbl(wRs.Fields("融資残高").Value), _
                P8.FCDbl(wRs.Fields("支払額").Value), _
                P8.FCDbl(wRs.Fields("洗替").Value), _
                P8.FCDbl(wRs.Fields("計上").Value), _
                P8.FCDbl(wRs.Fields("支払利息").Value)
    
        wRs.MoveNext
        Loop
        End If
        wRs.Close
        Set wRs = Nothing
            
        Close #1 '出力ファイルを閉じる
    
    End If

'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

Private Sub MX040_支払利息推移表_杉村倉庫(pCsvFileName As String)
'
    Dim ws01 As String, wstr2 As String
    Dim w開始年月日 As Date
    Dim w推移表区分 As String
    Dim w番号 As String
    Dim wOrder As String
    Dim ws銀行番号 As String, ws銀行名 As String
    Dim ws利息区分 As String, ws利息区分名 As String
    Dim WKGCNT As Integer
    Dim wc_融資金額 As Currency, wc_合計 As Currency, wc_支払利息1 As Currency, wc_支払利息2 As Currency, wc_支払利息3 As Currency, wc_支払利息4 As Currency
    Dim WKGCNT1 As Integer
    Dim wc_融資金額1 As Currency, wc_合計1 As Currency, wc_支払利息11 As Currency, wc_支払利息12 As Currency, wc_支払利息13 As Currency, wc_支払利息14 As Currency
    Dim j As Integer
    Dim ws支払利息1 As String, ws支払利息2 As String, ws支払利息3 As String, ws支払利息4 As String
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "西暦") '年度変換後
    w推移表タイトル = MUA010_推移表ファイル作成("", "", w開始年月日, GRpt.推移, 12)

    w番号 = Right("00" + CStr(GInt1), 2)
'
    wstr = "Select "
    wstr = wstr & "K.借入番号,"
    wstr = wstr & "K.利息区分,"
    wstr = wstr & "K.銀行番号,"
    wstr = wstr & "K.借入金種別区分,"
    wstr = wstr & "IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As 利息区分名,"
    wstr = wstr & "G.銀行名,"
    wstr = wstr & "S.借入金種別名,"
    
    wstr = wstr & "利率_" + w番号 + " As 利率,"
    
    wstr = wstr & "K.融資金額 As 融資金額,"
    
    '杉村倉庫仕様
    For j = 1 To 4
        w番号 = Right("00" + CStr(j), 2)
        wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & ",Z.未払利息増_" & w番号 & ") As 支払利息" + w番号 + ","
    Next
    
    '合計
    w番号 = "01"
    wstr = wstr & "(IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & ",Z.未払利息増_" & w番号 & ")"
    For j = 2 To GInt1
        w番号 = Right("00" + CStr(j), 2)
        wstr = wstr & " + IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & ",Z.未払利息増_" & w番号 & ")"
    Next
    wstr = wstr & ") As 合計"
    
    wstr = wstr & " FROM ((((DCDA010_借入残高推移表結果 As Z"
    wstr = wstr & " INNER JOIN DCDA010_借入残高推移表結果２ As Z2"
    wstr = wstr & " ON Z.借入番号=Z2.借入番号)"
    wstr = wstr & " INNER JOIN DCIA010_借入金ワーク As K"
    wstr = wstr & " ON Z.借入番号=K.借入番号)"
    wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr & " ON K.銀行番号=G.銀行番号)"
    wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
    wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分)"
    
    'Where 条件
    'All 0 は表示しない
    wstr = wstr & " Where ("
    '融資
    wstr = wstr & " Z.融資_01<>0"
    For j = 2 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.融資_" & ws01 & "<>0"
    Next j
    '元金
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.元金_" & ws01 & "<>0"
    Next j
    '利息
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.利息_" & ws01 & "<>0"
    Next j
    '返済
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.返済_" & ws01 & "<>0"
    Next j
    '解約
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.解約_" & ws01 & "<>0"
    Next j
    '保証
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.保証_" & ws01 & "<>0"
    Next j
    '残高
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.残高_" & ws01 & "<>0"
    Next j
    '前払利息
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.前払利息_" & ws01 & "<>0"
    Next j
    '前払利息増
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.前払利息増_" & ws01 & "<>0"
    Next j
    '前払利息減
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.前払利息減_" & ws01 & "<>0"
    Next j
    '未払利息
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.未払利息_" & ws01 & "<>0"
    Next j
    '未払利息増
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.未払利息増_" & ws01 & "<>0"
    Next j
    '未払利息減
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z.未払利息減_" & ws01 & "<>0"
    Next j
    '損益利息額
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr = wstr & " Or Z2.損益利息額_" & ws01 & "<>0"
    Next j
    wstr = wstr & " )"
    
    wstr = wstr & " Order by K.銀行番号,K.利息区分,K.借入番号"
    
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        
        '名称
        Write #1, _
            "借入番号", "銀行番号", "銀行名", "利息区分名", "利率/件数", "融資金額", "合計", _
            w推移表タイトル.X番目年月(1), w推移表タイトル.X番目年月(2), w推移表タイトル.X番目年月(3), w推移表タイトル.X番目年月(4)
    
        ws銀行番号 = "": ws銀行名 = "": ws利息区分 = "": ws利息区分名 = ""
        
        Do Until wRs.eof
        
            '利息区分計
            If ws銀行番号 & ws利息区分 <> "" _
            And ws銀行番号 & ws利息区分 <> P8.FCStr(wRs.Fields("銀行番号").Value) & P8.FCStr(wRs.Fields("利息区分").Value) Then
                
                ws支払利息1 = "": ws支払利息2 = "": ws支払利息3 = "": ws支払利息4 = ""
                ws支払利息1 = wc_支払利息11
                If GInt1 = 2 Then
                    ws支払利息2 = wc_支払利息12
                ElseIf GInt1 = 3 Then
                    ws支払利息2 = wc_支払利息12
                    ws支払利息3 = wc_支払利息13
                ElseIf GInt1 = 4 Then
                    ws支払利息2 = wc_支払利息12
                    ws支払利息3 = wc_支払利息13
                    ws支払利息4 = wc_支払利息14
                End If
                
                Write #1, _
                    ws利息区分名 & " 計", _
                    ws銀行番号, _
                    ws銀行名, _
                    ws利息区分名, _
                    WKGCNT1, _
                    wc_融資金額1, _
                    wc_合計1, _
                    ws支払利息1, _
                    ws支払利息2, _
                    ws支払利息3, _
                    ws支払利息4
                            
                    WKGCNT1 = 0
                    wc_融資金額1 = 0
                    wc_合計1 = 0
                    wc_支払利息11 = 0
                    wc_支払利息12 = 0
                    wc_支払利息13 = 0
                    wc_支払利息14 = 0
            End If
            
            '銀行計
            If ws銀行番号 <> "" And ws銀行番号 <> P8.FCStr(wRs.Fields("銀行番号").Value) Then
            
                ws支払利息1 = "": ws支払利息2 = "": ws支払利息3 = "": ws支払利息4 = ""
                ws支払利息1 = wc_支払利息1
                If GInt1 = 2 Then
                    ws支払利息2 = wc_支払利息2
                ElseIf GInt1 = 3 Then
                    ws支払利息2 = wc_支払利息2
                    ws支払利息3 = wc_支払利息3
                ElseIf GInt1 = 4 Then
                    ws支払利息2 = wc_支払利息2
                    ws支払利息3 = wc_支払利息3
                    ws支払利息4 = wc_支払利息4
                End If
                
                Write #1, _
                    "小計", _
                    ws銀行番号, _
                    ws銀行名, _
                    "", _
                    WKGCNT, _
                    wc_融資金額, _
                    wc_合計, _
                    ws支払利息1, _
                    ws支払利息2, _
                    ws支払利息3, _
                    ws支払利息4
                            
                    WKGCNT = 0
                    wc_融資金額 = 0
                    wc_合計 = 0
                    wc_支払利息1 = 0
                    wc_支払利息2 = 0
                    wc_支払利息3 = 0
                    wc_支払利息4 = 0
            End If
            
            If GRpt.詳細表示 = 1 Then
                ws支払利息1 = "": ws支払利息2 = "": ws支払利息3 = "": ws支払利息4 = ""
                ws支払利息1 = P8.FCDbl(wRs.Fields("支払利息01").Value)
                If GInt1 = 2 Then
                    ws支払利息2 = P8.FCDbl(wRs.Fields("支払利息02").Value)
                ElseIf GInt1 = 3 Then
                    ws支払利息2 = P8.FCDbl(wRs.Fields("支払利息02").Value)
                    ws支払利息3 = P8.FCDbl(wRs.Fields("支払利息03").Value)
                ElseIf GInt1 = 4 Then
                    ws支払利息2 = P8.FCDbl(wRs.Fields("支払利息02").Value)
                    ws支払利息3 = P8.FCDbl(wRs.Fields("支払利息03").Value)
                    ws支払利息4 = P8.FCDbl(wRs.Fields("支払利息04").Value)
                End If
                
                Write #1, _
                    P8.FCStr(wRs.Fields("借入番号").Value), _
                    P8.FCStr(wRs.Fields("銀行番号").Value), _
                    P8.FCStr(wRs.Fields("銀行名").Value), _
                    P8.FCStr(wRs.Fields("利息区分名").Value), _
                    P8.FCDbl(wRs.Fields("利率").Value), _
                    P8.FCStr(wRs.Fields("融資金額").Value), _
                    P8.FCDbl(wRs.Fields("合計").Value), _
                    ws支払利息1, _
                    ws支払利息2, _
                    ws支払利息3, _
                    ws支払利息4
            End If
            
            ws銀行番号 = P8.FCStr(wRs.Fields("銀行番号").Value)
            ws銀行名 = P8.FCStr(wRs.Fields("銀行名").Value)
            ws利息区分 = P8.FCStr(wRs.Fields("利息区分").Value)
            ws利息区分名 = P8.FCStr(wRs.Fields("利息区分名").Value)
            
            '利息区分計
            WKGCNT1 = WKGCNT1 + 1
            wc_融資金額1 = wc_融資金額1 + wRs.Fields("融資金額").Value
            wc_合計1 = wc_合計1 + wRs.Fields("合計").Value
            wc_支払利息11 = wc_支払利息11 + wRs.Fields("支払利息01").Value
            wc_支払利息12 = wc_支払利息12 + wRs.Fields("支払利息02").Value
            wc_支払利息13 = wc_支払利息13 + wRs.Fields("支払利息03").Value
            wc_支払利息14 = wc_支払利息14 + wRs.Fields("支払利息04").Value
            
            '銀行計
            WKGCNT = WKGCNT + 1
            wc_融資金額 = wc_融資金額 + wRs.Fields("融資金額").Value
            wc_合計 = wc_合計 + wRs.Fields("合計").Value
            wc_支払利息1 = wc_支払利息1 + wRs.Fields("支払利息01").Value
            wc_支払利息2 = wc_支払利息2 + wRs.Fields("支払利息02").Value
            wc_支払利息3 = wc_支払利息3 + wRs.Fields("支払利息03").Value
            wc_支払利息4 = wc_支払利息4 + wRs.Fields("支払利息04").Value
            
        wRs.MoveNext
        Loop
        
        If ws銀行番号 <> "" Then
           '利息区分計
            ws支払利息1 = "": ws支払利息2 = "": ws支払利息3 = "": ws支払利息4 = ""
            ws支払利息1 = wc_支払利息11
            If GInt1 = 2 Then
                ws支払利息2 = wc_支払利息12
            ElseIf GInt1 = 3 Then
                ws支払利息2 = wc_支払利息12
                ws支払利息3 = wc_支払利息13
            ElseIf GInt1 = 4 Then
                ws支払利息2 = wc_支払利息12
                ws支払利息3 = wc_支払利息13
                ws支払利息4 = wc_支払利息14
            End If
            
            Write #1, _
                ws利息区分名 & " 計", _
                ws銀行番号, _
                ws銀行名, _
                ws利息区分名, _
                WKGCNT1, _
                wc_融資金額1, _
                wc_合計1, _
                ws支払利息1, _
                ws支払利息2, _
                ws支払利息3, _
                ws支払利息4
       
           '銀行計
            ws支払利息1 = "": ws支払利息2 = "": ws支払利息3 = "": ws支払利息4 = ""
            ws支払利息1 = wc_支払利息1
            If GInt1 = 2 Then
                ws支払利息2 = wc_支払利息2
            ElseIf GInt1 = 3 Then
                ws支払利息2 = wc_支払利息2
                ws支払利息3 = wc_支払利息3
            ElseIf GInt1 = 4 Then
                ws支払利息2 = wc_支払利息2
                ws支払利息3 = wc_支払利息3
                ws支払利息4 = wc_支払利息4
            End If
            
            Write #1, _
                "小計", _
                ws銀行番号, _
                ws銀行名, _
                "", _
                WKGCNT, _
                wc_融資金額, _
                wc_合計, _
                ws支払利息1, _
                ws支払利息2, _
                ws支払利息3, _
                ws支払利息4
        End If
    
    End If
    wRs.Close
    Set wRs = Nothing
    
    '合計
    wstr2 = "select "
    wstr2 = wstr2 & "count(Z.借入番号) As 件数,"
    wstr2 = wstr2 & "SUM(K.融資金額) As 融資金額,"
    '杉村倉庫仕様
    For j = 1 To 4
        w番号 = Right("00" + CStr(j), 2)
        wstr2 = wstr2 & "SUM(IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & ",Z.未払利息増_" & w番号 & ")) As 支払利息" + w番号 + ","
    Next
    
    '合計
    w番号 = "01"
    wstr2 = wstr2 & "(SUM(IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & ",Z.未払利息増_" & w番号 & "))"
    For j = 2 To GInt1
        w番号 = Right("00" + CStr(j), 2)
        wstr2 = wstr2 & " + SUM(IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & ",Z.未払利息増_" & w番号 & "))"
    Next
    wstr2 = wstr2 & ") As 合計"
    
    wstr2 = wstr2 & " FROM ((((DCDA010_借入残高推移表結果 As Z"
    wstr2 = wstr2 & " INNER JOIN DCDA010_借入残高推移表結果２ As Z2"
    wstr2 = wstr2 & " ON Z.借入番号=Z2.借入番号)"
    wstr2 = wstr2 & " INNER JOIN DCIA010_借入金ワーク As K"
    wstr2 = wstr2 & " ON Z.借入番号=K.借入番号)"
    wstr2 = wstr2 & " INNER JOIN DAAA040_銀行マスタ As G"
    wstr2 = wstr2 & " ON K.銀行番号=G.銀行番号)"
    wstr2 = wstr2 & " LEFT JOIN DAAA116_借入金種別 As S"
    wstr2 = wstr2 & " ON K.借入金種別区分 = S.借入金種別区分)"
    
    'Where 条件
    'All 0 は表示しない
    wstr2 = wstr2 & " Where ("
    '融資
    wstr2 = wstr2 & " Z.融資_01<>0"
    For j = 2 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr2 = wstr2 & " Or Z.融資_" & ws01 & "<>0"
    Next j
    '元金
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr2 = wstr2 & " Or Z.元金_" & ws01 & "<>0"
    Next j
    '利息
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr2 = wstr2 & " Or Z.利息_" & ws01 & "<>0"
    Next j
    '返済
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr2 = wstr2 & " Or Z.返済_" & ws01 & "<>0"
    Next j
    '解約
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr2 = wstr2 & " Or Z.解約_" & ws01 & "<>0"
    Next j
    '保証
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr2 = wstr2 & " Or Z.保証_" & ws01 & "<>0"
    Next j
    '残高
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr2 = wstr2 & " Or Z.残高_" & ws01 & "<>0"
    Next j
    '前払利息
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr2 = wstr2 & " Or Z.前払利息_" & ws01 & "<>0"
    Next j
    '前払利息増
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr2 = wstr2 & " Or Z.前払利息増_" & ws01 & "<>0"
    Next j
    '前払利息減
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr2 = wstr2 & " Or Z.前払利息減_" & ws01 & "<>0"
    Next j
    '未払利息
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr2 = wstr2 & " Or Z.未払利息_" & ws01 & "<>0"
    Next j
    '未払利息増
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr2 = wstr2 & " Or Z.未払利息増_" & ws01 & "<>0"
    Next j
    '未払利息減
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr2 = wstr2 & " Or Z.未払利息減_" & ws01 & "<>0"
    Next j
    '損益利息額
    For j = 1 To GInt1
        ws01 = Right("00" + CStr(j), 2)
        wstr2 = wstr2 & " Or Z2.損益利息額_" & ws01 & "<>0"
    Next j
    wstr2 = wstr2 & " )"
    
    Call AdoRecordsetOpen(GDb, wRs, wstr2)
    If Not wRs.eof Then
    Do Until wRs.eof
    
        ws支払利息1 = "": ws支払利息2 = "": ws支払利息3 = "": ws支払利息4 = ""
        ws支払利息1 = P8.FCDbl(wRs.Fields("支払利息01").Value)
        If GInt1 = 2 Then
            ws支払利息2 = P8.FCDbl(wRs.Fields("支払利息02").Value)
        ElseIf GInt1 = 3 Then
            ws支払利息2 = P8.FCDbl(wRs.Fields("支払利息02").Value)
            ws支払利息3 = P8.FCDbl(wRs.Fields("支払利息03").Value)
        ElseIf GInt1 = 4 Then
            ws支払利息2 = P8.FCDbl(wRs.Fields("支払利息02").Value)
            ws支払利息3 = P8.FCDbl(wRs.Fields("支払利息03").Value)
            ws支払利息4 = P8.FCDbl(wRs.Fields("支払利息04").Value)
        End If
            
        Write #1, _
            "総合計", "", "", "", wRs.Fields("件数").Value, _
            P8.FCDbl(wRs.Fields("融資金額").Value), _
            P8.FCDbl(wRs.Fields("合計").Value), _
            ws支払利息1, _
            ws支払利息2, _
            ws支払利息3, _
            ws支払利息4

    wRs.MoveNext
    Loop
    End If
    wRs.Close
    Set wRs = Nothing
        
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_長期プライムレート
'------------------------------------------------
Private Sub MX040_長期プライムレート(pCsvFileName As String)
'
    Dim ws01 As String
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wstr = ""
    wstr = wstr & "Select "
    wstr = wstr & "Format(T.年月日,'yyyy/mm/dd') As 年月日,"
    wstr = wstr & "Format(T.長期プライムレート,'#,##0.00000') As 基準金利レート"
    wstr = wstr & " FROM DBDA010_借入金長期プライムレート As T"
    wstr = wstr & " Where 基準金利区分='" & GRpt.テキスト_01 & "'"
    wstr = wstr & " Order BY 年月日"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    If wRs.RecordCount = 0 Then
        MsgBox "出力データがありません", vbInformation
        
        wRs.Close
        Set wRs = Nothing
    End If
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        '名称
        Write #1, _
            "年月日", "基準金利レート"
            
    Do Until wRs.eof
        Write #1, _
            P8.FCStr(wRs.Fields("年月日").Value), _
            P8.FCStr(wRs.Fields("基準金利レート").Value)
    wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
    
    MsgBox "出力しました", vbInformation
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_借入明細表
'------------------------------------------------
Public Sub MX040_借入金時価評価明細表(pCsvFileName As String)
'
    Dim ws01 As String
    Dim wsNendo As String
    Dim wd01_1 As Double, wd01_2 As Double, wd01_3 As Double
    Dim wd02_1 As Double, wd02_2 As Double, wd02_3 As Double
    Dim wv返済年月日 As Variant
    
    Dim wd決算日1 As Date, wd決算日2 As Date
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    wd決算日1 = P8.FCDate(GRpt.テキスト_01)
    wd決算日2 = C年月日.GetDate("設定", DateAdd("yyyy", 1, wd決算日1), G基本情報.決算締日)
    
    DoEvents
'
    wstr = "SELECT "
    GRpt.帳票名 = "借入金時価評価明細表"
    wstr = wstr & "M.借入番号,"
    wstr = wstr & "Format(M.返済年月日,'" & Gfmtcsv年月日 & "') As 返済年月日,"
    wstr = wstr + "Format(M.返済年月日,'" & Gfmt年月日 & "') As w返済年月日,"
    wstr = wstr & "Format(M.利息計算年月日,'" & Gfmtcsv年月日 & "') As 利息計算年月日,"
    wstr = wstr & "M.元金額,"
    wstr = wstr & "M.利息額,"
    wstr = wstr & "M.返済金額,"
    wstr = wstr & "M.融資残高,"
    wstr = wstr & "M.日割日数,"
    wstr = wstr & "format(M.指数,'#,##0.000000') As 指数,"
    wstr = wstr & "M.分母,"
    wstr = wstr & "M.現価係数,"
    wstr = wstr & "M.現在価値"
    wstr = wstr & " FROM DCDA010_借入金時価評価明細 As M"
    wstr = wstr & " INNER JOIN DBDA010_借入金 As K"
    wstr = wstr & " ON M.借入番号 = K.借入番号"
    wstr = wstr & " WHERE M.借入番号='" & GRpt.コンボ_01 & "'"
    wstr = wstr & " Order by M.返済年月日"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
        
    On Error GoTo Err_Hundle
            
        Write #1, _
                "借入番号", _
                "返済年月日", "利息計算年月日", "元金額", "利息額", "返済金額", _
                "融資残高", "日割日数", "現価係数", "時価評価額"
        
'        Write #1, _
'                "借入番号", _
'                "返済年月日", "利息計算年月日", "元金額", "利息額", "返済金額", _
'                "融資残高", "日割日数", "指数", "分母", _
'                "時価評価額"
    
        Do Until wRs.eof
        
            wv返済年月日 = C年月日.平成To西暦("年月日", P8.FCStr(wRs.Fields("w返済年月日").Value))
            If wd決算日1 < P8.FCDate(wv返済年月日) And P8.FCDate(wv返済年月日) <= wd決算日2 Then
                wd01_1 = Format(wd01_1 + P8.FCDbl(wRs.Fields("元金額").Value), "#,##0.00")
                wd02_1 = Format(wd02_1 + P8.FCDbl(wRs.Fields("現在価値").Value), "#,##0.00")
            ElseIf wd決算日2 < P8.FCDate(wv返済年月日) Then
                wd01_2 = Format(wd01_2 + P8.FCDbl(wRs.Fields("元金額").Value), "#,##0.00")
                wd02_2 = Format(wd02_2 + P8.FCDbl(wRs.Fields("現在価値").Value), "#,##0.00")
            End If
            
            wd01_3 = Format(wd01_3 + P8.FCDbl(wRs.Fields("元金額").Value), "#,##0.00")
            wd02_3 = Format(wd02_3 + P8.FCDbl(wRs.Fields("現在価値").Value), "#,##0.00")

            Write #1, _
                CStr(wRs.Fields("借入番号").Value), _
                CStr(wRs.Fields("返済年月日").Value), _
                CStr(wRs.Fields("利息計算年月日").Value), _
                CStr(wRs.Fields("元金額").Value), _
                CStr(wRs.Fields("利息額").Value), _
                CStr(wRs.Fields("返済金額").Value), _
                CStr(wRs.Fields("融資残高").Value), _
                CStr(wRs.Fields("日割日数").Value), _
                CStr(wRs.Fields("現価係数").Value), _
                CStr(wRs.Fields("現在価値").Value)

'            Write #1, _
'                CStr(wRs.Fields("借入番号").Value), _
'                CStr(wRs.Fields("返済年月日").Value), _
'                CStr(wRs.Fields("利息計算年月日").Value), _
'                CStr(wRs.Fields("元金額").Value), _
'                CStr(wRs.Fields("利息額").Value), _
'                CStr(wRs.Fields("返済金額").Value), _
'                CStr(wRs.Fields("融資残高").Value), _
'                CStr(wRs.Fields("日割日数").Value), _
'                CStr(wRs.Fields("指数").Value), _
'                CStr(wRs.Fields("分母").Value), _
'                CStr(wRs.Fields("現在価値").Value)
                
            wRs.MoveNext
        Loop
        
            Write #1, _
                "", "", "", "", "", "", "", "", "", "", ""
                
            Write #1, _
                "", "", "", "元金額", "時価評価額", "時価評価損益", "", "", "", "", ""

            Write #1, _
                "", "", "1年内返済", wd01_1, wd02_1, wd02_1 - wd01_1, "", "", "", "", ""
                
            Write #1, _
                "", "", "1年超返済", wd01_2, wd02_2, wd02_2 - wd01_2, "", "", "", "", ""

            Write #1, _
                "", "", "合計", wd01_3, wd02_3, wd02_3 - wd01_3, "", "", "", "", ""
        
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
'
'    GDb.Execute wstr
'
'    DoEvents
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_借入金時価評価適用金利一覧
'------------------------------------------------
Private Sub MX040_借入金時価評価適用金利一覧(pCsvFileName As String)
'
    Dim j As Integer
    Dim wd01 As Double
    Dim wdzan(2) As Double
    Dim wiCnt(2) As Integer
    Dim ws01 As String, wsNendo As String
    Dim wsTbl As String, wstr2 As String
    
    Dim wsF01 As String, wsF02 As String
    
    Dim ws計 As String
    Dim ws銀行番号 As String, ws銀行名 As String
    Dim ws基準金利区分 As String, ws基準金利名 As String
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wsTbl = "DCDA010_借入金時価評価適用金利"
    If GRpt.作業 = "前期末" Then
        wsTbl = "DCDA010_借入金時価評価適用金利前期末"
    End If

    wstr = "SELECT "
    wstr = wstr & "G.銀行名 As 銀行名,"
    wstr = wstr & "S.基準金利名 As 基準金利名,"
    
    wstr = wstr & "KZ.銀行番号 As 銀行番号,"
    wstr = wstr & "K.基準金利区分 As 基準金利区分,"
    wstr = wstr & "KZ.借入番号 As 借入番号,"
    wstr = wstr & "Format(KZ.実行日,'" & Gfmtcsv年月日 & "') As 実行日,"
    wstr = wstr & "KZ.融資金額 As 融資金額,"
    wstr = wstr & "Format(KZ.利率,'#,##0.00000') As 利率,"
    wstr = wstr & "Format(KZ.最終返済実行日,'" & Gfmtcsv年月日 & "') As 最終返済年月日,"
    wstr = wstr & "KZ.決算時融資残高 As 期末残高,"
    wstr = wstr & "Format(KZ.借入時長期プライムレート, '#,##0.00000') As 借入時基準金利,"
    wstr = wstr & "Format(KZ.借入時プレミアム, '#,##0.00000') As 借入時プレミアム,"
    wstr = wstr & "Format(KZ.決算時長期プライムレート, '#,##0.00000') As 決算時基準金利,"
    wstr = wstr & "Format(KZ.時価評価適用プレミアム, '#,##0.00000') As 時価評価適用プレミアム,"
    wstr = wstr & "Format(KZ.決算時時価評価適用金利, '#,##0.00000') As 決算時時価評価適用金利"
    
    wstr = wstr & " From ((" & wsTbl & "  As KZ "
    wstr = wstr & " Inner JOIN DBDA010_借入金 As K"
    wstr = wstr & " ON KZ.借入番号 = K.借入番号)"
    wstr = wstr & " Inner JOIN DAAA040_銀行マスタ As G"
    wstr = wstr & " ON K.銀行番号 = G.銀行番号)"
    wstr = wstr & " LEFT JOIN DAAA116_基準金利 As S"
    wstr = wstr & " ON K.基準金利区分 = S.基準金利区分"
    
    'wstr = wstr & " Order By K.借入金種別区分,KZ.銀行番号,K.借入番号"
    wstr = wstr & " Order By K.基準金利区分,KZ.銀行番号,K.実行日"
    
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        
    '名称
    Write #1, _
        "借入番号", "基準金利区分", "基準金利名", "銀行番号", "銀行名", _
        "実行日", "融資金額", "利率", "最終返済年月日", "期末残高", _
        "借入時基準金利", "借入時プレミアム", _
        "決算時基準金利", "時価評価適用プレミアム", _
        "決算時時価評価適用金利"

    If Not wRs.eof Then
        Do Until wRs.eof
        
            wd01 = P8.FCDbl(wRs.Fields("期末残高").Value)
    
            If ws計 = "" Then
                ws計 = P8.FCStr(wRs.Fields("基準金利区分").Value) & P8.FCStr(wRs.Fields("銀行番号").Value)
                ws基準金利区分 = P8.FCStr(wRs.Fields("基準金利区分").Value)
                ws基準金利名 = P8.FCStr(wRs.Fields("基準金利名").Value)
                ws銀行番号 = P8.FCStr(wRs.Fields("銀行番号").Value)
                ws銀行名 = P8.FCStr(wRs.Fields("銀行名").Value)
            End If
        
            If ws計 <> "" And ws計 <> P8.FCStr(wRs.Fields("基準金利区分").Value) & P8.FCStr(wRs.Fields("銀行番号").Value) Then
            
                If ws基準金利区分 <> P8.FCStr(wRs.Fields("基準金利区分").Value) Then
                    '小計
                    Write #1, _
                        "小計", ws基準金利区分, ws基準金利名, ws銀行番号, ws銀行名, _
                        wiCnt(0), "", "", "", _
                        wdzan(0), "", "", "", "", ""
                        
                    '基準金利区分
                    Write #1, _
                        "基準金利区分計", ws基準金利区分, ws基準金利名, ws銀行番号, ws銀行名, _
                        wiCnt(1), "", "", "", _
                        wdzan(1), "", "", "", "", ""
                            
                    ws銀行番号 = P8.FCStr(wRs.Fields("銀行番号").Value)
                    ws銀行名 = P8.FCStr(wRs.Fields("銀行名").Value)
                    
                    wiCnt(0) = 0
                    wdzan(0) = 0
                    
                    ws基準金利区分 = P8.FCStr(wRs.Fields("基準金利区分").Value)
                    ws基準金利名 = P8.FCStr(wRs.Fields("基準金利名").Value)
                    
                    wiCnt(1) = 0
                    wdzan(1) = 0
                                    
                ElseIf ws銀行番号 <> "" And ws銀行番号 <> P8.FCStr(wRs.Fields("銀行番号").Value) Then
                        '小計
                        Write #1, _
                            "小計", ws基準金利区分, ws基準金利名, ws銀行番号, ws銀行名, _
                            wiCnt(0), "", "", "", _
                            wdzan(0), "", "", "", "", ""
                            
                        ws銀行番号 = P8.FCStr(wRs.Fields("銀行番号").Value)
                        ws銀行名 = P8.FCStr(wRs.Fields("銀行名").Value)
                        
                        wiCnt(0) = 0
                        wdzan(0) = 0
                End If
            
            End If
                    
                Write #1, _
                    P8.FCStr(wRs.Fields("借入番号").Value), _
                    P8.FCStr(wRs.Fields("基準金利区分").Value), _
                    P8.FCStr(wRs.Fields("基準金利名").Value), _
                    P8.FCStr(wRs.Fields("銀行番号").Value), _
                    P8.FCStr(wRs.Fields("銀行名").Value), _
                    P8.FCStr(wRs.Fields("実行日").Value), _
                    P8.FCDbl(wRs.Fields("融資金額").Value), _
                    P8.FCDbl(wRs.Fields("利率").Value), _
                    P8.FCStr(wRs.Fields("最終返済年月日").Value), _
                    wd01, _
                    P8.FCDbl(wRs.Fields("借入時基準金利").Value), _
                    P8.FCDbl(wRs.Fields("借入時プレミアム").Value), _
                    P8.FCDbl(wRs.Fields("決算時基準金利").Value), _
                    P8.FCDbl(wRs.Fields("時価評価適用プレミアム").Value), _
                    P8.FCDbl(wRs.Fields("決算時時価評価適用金利").Value)
                
                wiCnt(0) = wiCnt(0) + 1
                wdzan(0) = wdzan(0) + wd01
                
                wiCnt(1) = wiCnt(1) + 1
                wdzan(1) = wdzan(1) + wd01
                
                wiCnt(2) = wiCnt(2) + 1
                wdzan(2) = wdzan(2) + wd01
    
        wRs.MoveNext
        Loop
            
            '小計
            Write #1, _
                "小計", ws基準金利区分, ws基準金利名, ws銀行番号, ws銀行名, _
                wiCnt(0), "", "", "", _
                wdzan(0), "", "", "", "", ""
                
            '基準金利区分
            Write #1, _
                "基準金利区分計", ws基準金利区分, ws基準金利名, ws銀行番号, ws銀行名, _
                wiCnt(1), "", "", "", _
                wdzan(1), "", "", "", "", ""
    
            Write #1, _
                "長期借入金合計", "", "", "", "", _
                wiCnt(2), "", "", "", _
                wdzan(2), "", "", "", "", ""

    End If
    wRs.Close
    Set wRs = Nothing
            
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MX040_借入金時価評価一覧表
'------------------------------------------------
Private Sub MX040_借入金時価評価一覧表(pCsvFileName As String)
'
    Dim ws01 As String
    Dim wd01(3) As Double, wd02(3) As Double, wd03(3) As Double
    Dim wd04(3) As Double, wd05(3) As Double, wd06(3) As Double
    Dim wd07(3) As Double, wd08(3) As Double, wd09(3) As Double
    Dim wiCnt(3) As Integer
    Dim wsNendo As String, wsTbl As String
    
    Dim ws計 As String
    Dim ws銀行番号 As String, ws銀行名 As String
    Dim ws基準金利区分 As String, ws基準金利名 As String
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wsTbl = "DCDA010_借入金時価評価"
    If GRpt.作業 = "前期末" Then
        wsTbl = "DCDA010_借入金時価評価前期末"
    ElseIf GRpt.作業 = "比較" Then
        wsTbl = "DCDA010_借入金時価評価前期末比較増減"
    End If
'
    If GRpt.作業 <> "比較" Then
        wstr = "SELECT"
        wstr = wstr + " JZ.借入番号 As 借入番号"
        wstr = wstr + ",JZ.銀行番号 AS 銀行番号"
        wstr = wstr + ",S.基準金利名 As 基準金利名"
        wstr = wstr + ",K.基準金利区分 As 基準金利区分"
        wstr = wstr + ",G.銀行名 As 銀行名"
        wstr = wstr + ",Format(JZ.実行日,'" & Gfmtcsv年月日 & "') AS 実行日"
        wstr = wstr + ",JZ.融資金額 AS 融資金額"
        wstr = wstr + ",Format(JZ.利率,'#,##0.00000') AS 利率"
        wstr = wstr + ",Format(JZ.最終返済実行日,'" & Gfmtcsv年月日 & "') AS 最終返済実行日"
        wstr = wstr + ",JZ.決算年月日 AS 決算年月日"
        wstr = wstr + ",JZ.合計決算時融資残高 AS 期末残高"
        wstr = wstr + ",JZ.合計時価評価額 AS 時価評価額"
        wstr = wstr + ",JZ.合計時価損益 AS 時価評価損益"
        wstr = wstr + ",JZ.年以内返済予定元金 AS 長期借入金2"
        wstr = wstr + ",JZ.年以内返済予定時価評価額 AS 時価評価額2"
        wstr = wstr + ",JZ.年以内返済予定時価損益 AS 時価評価損益2"
        wstr = wstr + ",JZ.年超返済予定元金 AS 長期借入金3"
        wstr = wstr + ",JZ.年超返済予定時価評価額 AS 時価評価額3"
        wstr = wstr + ",JZ.年超返済予定時価損益 AS 時価評価損益3"
    
        wstr = wstr + " FROM ((" & wsTbl & " AS JZ"
        wstr = wstr + " INNER JOIN DBDA010_借入金 As K"
        wstr = wstr + " ON JZ.借入番号 = K.借入番号)"
        wstr = wstr + " LEFT JOIN DAAA116_基準金利 As S"
        wstr = wstr + " ON K.基準金利区分 = S.基準金利区分)"
        wstr = wstr + " LEFT JOIN DAAA040_銀行マスタ As G"
        wstr = wstr + " ON K.銀行番号 = G.銀行番号"
    
        'wstr = wstr & " Order by K.借入金種別区分,K.銀行番号,K.借入番号"
        wstr = wstr & " Order by K.基準金利区分,K.銀行番号,K.実行日"
    
    Else
        wstr = "SELECT"
        wstr = wstr + " JZ.借入番号 As 借入番号"
        wstr = wstr + ",First(JZ.銀行番号) AS 銀行番号"
        wstr = wstr + ",First(S.基準金利名) As 基準金利名"
        wstr = wstr + ",First(K.基準金利区分) As 基準金利区分"
        wstr = wstr + ",First(G.銀行名) As 銀行名"
        wstr = wstr + ",First(Format(JZ.実行日,'" & Gfmtcsv年月日 & "')) AS 実行日"
        wstr = wstr + ",First(JZ.融資金額) AS 融資金額"
        wstr = wstr + ",Format(First(JZ.利率),'#,##0.00000') AS 利率"
        wstr = wstr + ",First(Format(JZ.最終返済実行日,'" & Gfmtcsv年月日 & "')) AS 最終返済実行日"
        wstr = wstr + ",First(JZ.決算年月日) AS 決算年月日"
        wstr = wstr + ",Sum(JZ.合計決算時融資残高) AS 期末残高"
        wstr = wstr + ",Sum(JZ.合計時価評価額) AS 時価評価額"
        wstr = wstr + ",Sum(JZ.合計時価損益) AS 時価評価損益"
        wstr = wstr + ",Sum(JZ.年以内返済予定元金) AS 長期借入金2"
        wstr = wstr + ",Sum(JZ.年以内返済予定時価評価額) AS 時価評価額2"
        wstr = wstr + ",Sum(JZ.年以内返済予定時価損益) AS 時価評価損益2"
        wstr = wstr + ",Sum(JZ.年超返済予定元金) AS 長期借入金3"
        wstr = wstr + ",Sum(JZ.年超返済予定時価評価額) AS 時価評価額3"
        wstr = wstr + ",Sum(JZ.年超返済予定時価損益) AS 時価評価損益3"
    
        wstr = wstr + " FROM ((" & wsTbl & " AS JZ"
        wstr = wstr + " INNER JOIN DBDA010_借入金 As K"
        wstr = wstr + " ON JZ.借入番号 = K.借入番号)"
        wstr = wstr + " LEFT JOIN DAAA116_基準金利 As S"
        wstr = wstr + " ON K.基準金利区分 = S.基準金利区分)"
        wstr = wstr + " LEFT JOIN DAAA040_銀行マスタ As G"
        wstr = wstr + " ON K.銀行番号 = G.銀行番号"
    
        wstr = wstr + " GROUP BY JZ.借入番号"
        'wstr = wstr + " ORDER BY First(K.借入金種別区分),First(JZ.銀行番号),JZ.借入番号"
        wstr = wstr + " ORDER BY First(K.基準金利区分),First(JZ.銀行番号),First(JZ.実行日)"
    
    End If
    
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
    
    '名称
    If GRpt.作業 = "比較" Then
        Write #1, _
            "借入番号", "基準金利区分", "基準金利名", "銀行番号", "銀行名", _
            "実行日", "融資金額", "利率", "最終返済年月日", _
            "期末残高増減額", "時価評価額増減額", "時価評価損益増減額", _
            "流動負債_長期借入金増減額", "流動負債_時価評価額増減額", "流動負債_時価評価損益増減額", _
            "固定負債_長期借入金増減額", "固定負債_時価評価額増減額", "固定負債_時価評価損益増減額"
    Else
        Write #1, _
            "借入番号", "基準金利区分", "基準金利名", "銀行番号", "銀行名", _
            "実行日", "融資金額", "利率", "最終返済年月日", _
            "期末残高", "時価評価額", "時価評価損益", _
            "流動負債_長期借入金", "流動負債_時価評価額", "流動負債_時価評価損益", _
            "固定負債_長期借入金", "固定負債_時価評価額", "固定負債_時価評価損益"
    End If

    If Not wRs.eof Then
        Do Until wRs.eof
        
            wd01(0) = P8.FCDbl(wRs.Fields("期末残高").Value)
            wd02(0) = P8.FCDbl(wRs.Fields("時価評価額").Value)
            wd03(0) = P8.FCDbl(wRs.Fields("時価評価損益").Value)
            wd04(0) = P8.FCDbl(wRs.Fields("長期借入金2").Value)
            wd05(0) = P8.FCDbl(wRs.Fields("時価評価額2").Value)
            wd06(0) = P8.FCDbl(wRs.Fields("時価評価損益2").Value)
            wd07(0) = P8.FCDbl(wRs.Fields("長期借入金3").Value)
            wd08(0) = P8.FCDbl(wRs.Fields("時価評価額3").Value)
            wd09(0) = P8.FCDbl(wRs.Fields("時価評価損益3").Value)
            
            If ws計 = "" Then
                ws計 = P8.FCStr(wRs.Fields("基準金利区分").Value) & P8.FCStr(wRs.Fields("銀行番号").Value)
                ws基準金利区分 = P8.FCStr(wRs.Fields("基準金利区分").Value)
                ws基準金利名 = P8.FCStr(wRs.Fields("基準金利名").Value)
                ws銀行番号 = P8.FCStr(wRs.Fields("銀行番号").Value)
                ws銀行名 = P8.FCStr(wRs.Fields("銀行名").Value)
            End If
        
            If ws計 <> "" And ws計 <> P8.FCStr(wRs.Fields("基準金利区分").Value) & P8.FCStr(wRs.Fields("銀行番号").Value) Then
            
                If ws基準金利区分 <> P8.FCStr(wRs.Fields("基準金利区分").Value) Then
                    '小計
                    Write #1, _
                        "小計", ws基準金利区分, ws基準金利名, ws銀行番号, ws銀行名, _
                        wiCnt(1), "", "", "", _
                        wd01(1), wd02(1), wd03(1), wd04(1), wd05(1), wd06(1), wd07(1), wd08(1), wd09(1)
                    
                    '基準金利区分計
                    Write #1, _
                        "基準金利計", ws基準金利区分, ws基準金利名, ws銀行番号, ws銀行名, _
                        wiCnt(2), "", "", "", _
                        wd01(2), wd02(2), wd03(2), wd04(2), wd05(2), wd06(2), wd07(2), wd08(2), wd09(2)
                        
                    ws銀行番号 = P8.FCStr(wRs.Fields("銀行番号").Value)
                    ws銀行名 = P8.FCStr(wRs.Fields("銀行名").Value)
                    
                    wiCnt(1) = 0
                    wd01(1) = 0
                    wd02(1) = 0
                    wd03(1) = 0
                    wd04(1) = 0
                    wd05(1) = 0
                    wd06(1) = 0
                    wd07(1) = 0
                    wd08(1) = 0
                    wd09(1) = 0
                        
                    ws基準金利区分 = P8.FCStr(wRs.Fields("基準金利区分").Value)
                    ws基準金利名 = P8.FCStr(wRs.Fields("基準金利名").Value)
                    
                    wiCnt(2) = 0
                    wd01(2) = 0
                    wd02(2) = 0
                    wd03(2) = 0
                    wd04(2) = 0
                    wd05(2) = 0
                    wd06(2) = 0
                    wd07(2) = 0
                    wd08(2) = 0
                    wd09(2) = 0
                
                ElseIf ws銀行番号 <> "" And ws銀行番号 <> P8.FCStr(wRs.Fields("銀行番号").Value) Then
                    '小計
                    Write #1, _
                        "小計", ws基準金利区分, ws基準金利名, ws銀行番号, ws銀行名, _
                        wiCnt(1), "", "", "", _
                        wd01(1), wd02(1), wd03(1), wd04(1), wd05(1), wd06(1), wd07(1), wd08(1), wd09(1)
                        
                    ws銀行番号 = P8.FCStr(wRs.Fields("銀行番号").Value)
                    ws銀行名 = P8.FCStr(wRs.Fields("銀行名").Value)
                    
                    wiCnt(1) = 0
                    wd01(1) = 0
                    wd02(1) = 0
                    wd03(1) = 0
                    wd04(1) = 0
                    wd05(1) = 0
                    wd06(1) = 0
                    wd07(1) = 0
                    wd08(1) = 0
                    wd09(1) = 0
                End If
            
            End If
            
            Write #1, _
                P8.FCStr(wRs.Fields("借入番号").Value), _
                P8.FCStr(wRs.Fields("基準金利区分").Value), _
                P8.FCStr(wRs.Fields("基準金利名").Value), _
                P8.FCStr(wRs.Fields("銀行番号").Value), _
                P8.FCStr(wRs.Fields("銀行名").Value), _
                P8.FCStr(wRs.Fields("実行日").Value), _
                P8.FCStr(wRs.Fields("融資金額").Value), _
                P8.FCStr(wRs.Fields("利率").Value), _
                P8.FCStr(wRs.Fields("最終返済実行日").Value), _
                wd01(0), wd02(0), wd03(0), wd04(0), wd05(0), wd06(0), wd07(0), wd08(0), wd09(0)
    
            '小計
            wiCnt(1) = wiCnt(1) + 1
            wd01(1) = wd01(1) + wd01(0)
            wd02(1) = wd02(1) + wd02(0)
            wd03(1) = wd03(1) + wd03(0)
            wd04(1) = wd04(1) + wd04(0)
            wd05(1) = wd05(1) + wd05(0)
            wd06(1) = wd06(1) + wd06(0)
            wd07(1) = wd07(1) + wd07(0)
            wd08(1) = wd08(1) + wd08(0)
            wd09(1) = wd09(1) + wd09(0)
            
            '基準金利区分計
            wiCnt(2) = wiCnt(2) + 1
            wd01(2) = wd01(2) + wd01(0)
            wd02(2) = wd02(2) + wd02(0)
            wd03(2) = wd03(2) + wd03(0)
            wd04(2) = wd04(2) + wd04(0)
            wd05(2) = wd05(2) + wd05(0)
            wd06(2) = wd06(2) + wd06(0)
            wd07(2) = wd07(2) + wd07(0)
            wd08(2) = wd08(2) + wd08(0)
            wd09(2) = wd09(2) + wd09(0)
            
            '合計
            wiCnt(3) = wiCnt(3) + 1
            wd01(3) = wd01(3) + wd01(0)
            wd02(3) = wd02(3) + wd02(0)
            wd03(3) = wd03(3) + wd03(0)
            wd04(3) = wd04(3) + wd04(0)
            wd05(3) = wd05(3) + wd05(0)
            wd06(3) = wd06(3) + wd06(0)
            wd07(3) = wd07(3) + wd07(0)
            wd08(3) = wd08(3) + wd08(0)
            wd09(3) = wd09(3) + wd09(0)
        
        wRs.MoveNext
        Loop
                
            '小計
            Write #1, _
                "小計", ws基準金利区分, ws基準金利名, ws銀行番号, ws銀行名, _
                wiCnt(1), "", "", "", _
                wd01(1), wd02(1), wd03(1), wd04(1), wd05(1), wd06(1), wd07(1), wd08(1), wd09(1)
    
            '基準金利区分計
            Write #1, _
                "基準金利計", ws基準金利区分, ws基準金利名, ws銀行番号, ws銀行名, _
                wiCnt(2), "", "", "", _
                wd01(2), wd02(2), wd03(2), wd04(2), wd05(2), wd06(2), wd07(2), wd08(2), wd09(2)
        
            Write #1, _
                "長期借入金合計", "", "", "", "", _
                wiCnt(3), "", "", "", _
                wd01(3), wd02(3), wd03(3), wd04(3), wd05(3), wd06(3), wd07(3), wd08(3), wd09(3)
    End If
    
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
    
    DoEvents
    
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub

'------------------------------------------------
' MXA040_基準金利レート取込
'------------------------------------------------
Public Function MXA040_基準金利レート取込(pCsvName As String, pKbn As String) As Boolean
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    
    Dim j As Long, k As Long
    Dim ws01 As String
    Dim wsName As String, wsValue As String
    Dim wsMsg As String
'
    MXA040_基準金利レート取込 = False
'
    On Error GoTo MXA040_基準金利レート取込_ERR
'
    '項目名で取り込むので名称セット
    '----------< 読込 >-------------------------------------------------------------
    Call MXA040_CsvInit
    '
    Call MXA040_CsvAdd("年月日")
    Call MXA040_CsvAdd("レート")
    
    wレコード = MXA040_CsvRead(pCsvName, 2) '1行目タイトル
'
    If UBound(wレコード) = 0 Then
        Exit Function
    End If
'
    For k = 1 To UBound(wレコード)
        For j = 1 To UBound(wレコード(k).xName)
            wsName = wレコード(k).xName(j)
            wsValue = CStr(wレコード(k).xValue(j))
            
            'Data Check
            If wsName = "年月日" Then
                ws01 = Format(wsValue, "yyyy/mm/dd")
                If Not IsDate(ws01) Then
                    GoTo MXA040_基準金利レート取込_ERR_CHECK
                End If
            ElseIf wsName = "レート" Then
                If Not IsNumeric(wsValue) And wsValue <> "" Then
                    GoTo MXA040_基準金利レート取込_ERR_CHECK
                End If
            
            Else
                If Not IsNumeric(wsValue) And wsValue <> "" Then
                    GoTo MXA040_基準金利レート取込_ERR_CHECK
                End If
            End If
        
        Next j
    Next k
'
    ' =========================================
    '               CSVデータCHECK
    ' =========================================
'
    ' =========================================
    '               Delete
    ' =========================================
    If UBound(wレコード) >= 1 Then
        wstr = "Delete * From DBDA010_借入金長期プライムレート"
        wstr = wstr & " Where 基準金利区分='" & pKbn & "'"
        GDb.Execute wstr
    End If
    
    DoEvents
'
    ' =========================================
    '               UPDATE
    ' =========================================
    wstr = ""
    wstr = wstr + "Select * From DBDA010_借入金長期プライムレート"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        For j = 1 To UBound(wレコード)
             
            wRs.AddNew
        
                wRs("基準金利区分") = pKbn
                wRs("年月日") = Format(wレコード(j).xValue(1), "yyyy/mm/dd")
                wRs("長期プライムレート") = P8.FCDbl(wレコード(j).xValue(2))
            
            wRs.Update
          
        Next
    
    wRs.Close
    Set wRs = Nothing
'
    On Error GoTo 0
'
    MXA040_基準金利レート取込 = True
'
Exit Function
'----------< ERROR >----------------------------------------------------------------
MXA040_基準金利レート取込_ERR_CHECK:
    wsMsg = CStr(k) & "行目 " & wsName & ":" & wsValue & " を確認してください。"
    MsgBox wsMsg, vbInformation
    
    Exit Function
'
MXA040_基準金利レート取込_ERR:
    Err.Clear
    Exit Function
End Function

'------------------------------------------------
' MXA040_補助科目取込
'------------------------------------------------
Public Function MXA040_補助科目取込(pCsvName As String) As Boolean
'
    Dim j As Long, k As Long, l As Long
    Dim lp1 As Long, lp2 As Long
    Dim wi01 As Integer
    Dim ws01 As String
    Dim wsName As String, wsValue As String
    Dim wsMsg As String

    Dim FLG_CSVCheck As Boolean
    Dim wiCnt As Integer
    Dim wGinko(99) As String
    Dim wKanjo(99) As String
    Dim wsKanjo1 As String, wsKanjo2 As String
    Dim wsGinko1 As String, wsGinko2 As String
'
    MXA040_補助科目取込 = False
'
    On Error GoTo MXA040_補助科目取込_ERR
'
    '項目名で取り込むので名称セット
    Call MXA040_補助科目項目名_SET(pCsvName)
'
    If UBound(wレコード) = 0 Then
        Exit Function
    End If
'
    For k = 1 To UBound(wレコード)
        For j = 1 To UBound(wレコード(k).xName)
            wsName = wレコード(k).xName(j)
            wsValue = CStr(wレコード(k).xValue(j))
            
            'Data Check
            If wsName = "勘定科目" Then
                If wsValue = "" Then
                    GoTo MXA040_補助科目取込_ERR_CHECK
                End If
            ElseIf wsName = "勘定科目名" Then
                If wsValue = "" Then
                    GoTo MXA040_補助科目取込_ERR_CHECK
                End If
            ElseIf wsName = "銀行番号" Then
                If wsValue = "" Then
                    GoTo MXA040_補助科目取込_ERR_CHECK
                End If
            ElseIf wsName = "銀行名" Then
'                If wsValue = "" Then
'                    GoTo MXA040_補助科目取込_ERR_CHECK
'                End If
            ElseIf wsName = "補助科目" Then
                If wsValue = "" Then
                    GoTo MXA040_補助科目取込_ERR_CHECK
                End If
            ElseIf wsName = "補助科目名" Then
                If wsValue = "" Then
                    GoTo MXA040_補助科目取込_ERR_CHECK
                End If
            
            Else
                If Not IsNumeric(wsValue) And wsValue <> "" Then
                    GoTo MXA040_補助科目取込_ERR_CHECK
                End If
            End If
        Next j
    Next k
'
    ' =========================================
    '           銀行番号
    ' =========================================
    wiCnt = 0
    wstr = ""
    wstr = wstr & "SELECT * From DAAA040_銀行マスタ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
      Do Until wRs.eof
        wGinko(wiCnt) = wRs("銀行番号")
        wiCnt = wiCnt + 1
        
        wRs.MoveNext
      Loop
    End If
'
    ' =========================================
    '           勘定科目
    ' =========================================
    wiCnt = 0
    wstr = ""
    wstr = wstr & "SELECT KJ1.借方勘定科目 AS 勘定科目"
    wstr = wstr & " FROM DABA010_勘定科目マスタ AS KJ1"
    wstr = wstr & " union SELECT KJ2.貸方勘定科目"
    wstr = wstr & " FROM DABA010_勘定科目マスタ AS KJ2"
    wstr = wstr & " GROUP BY KJ2.貸方勘定科目"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If Not wRs.eof Then
      Do Until wRs.eof
        wKanjo(wiCnt) = wRs("勘定科目")
        wiCnt = wiCnt + 1
      
        wRs.MoveNext
      Loop
    End If
'
    ' =========================================
    '           CSV CHECK
    ' =========================================
    For k = 1 To UBound(wレコード)
        For j = 1 To UBound(wレコード(k).xName)
            wsName = wレコード(k).xName(j)
            wsValue = CStr(wレコード(k).xValue(j))
            If wsName = "勘定科目" Then
                FLG_CSVCheck = False
                For l = 0 To 99
                    If wKanjo(l) = wsValue Then
                        FLG_CSVCheck = True
                        Exit For
                    End If
                    
                    If wKanjo(l) = "" Then
                        Exit For
                    End If
                Next l
            
                If FLG_CSVCheck = False Then
                    wsMsg = CStr(k) & "行目 勘定科目:" & wsValue & " を確認してください。"
                    MsgBox wsMsg, vbInformation

'                    Exit Function
                End If

            ElseIf wsName = "銀行番号" Then
                FLG_CSVCheck = False
                For l = 0 To 99
                    If wGinko(l) = wsValue Then
                        FLG_CSVCheck = True
                        Exit For
                    End If
                    
                    If wGinko(l) = "" Then
                        Exit For
                    End If
                Next l
            
                If FLG_CSVCheck = False Then
                    wsMsg = CStr(k) & "行目 銀行番号:" & wsValue & " を確認してください。"
                    MsgBox wsMsg, vbInformation

'                    Exit Function
                End If
            End If
        Next j
    Next k
'
    '重複CHECK
    For k = 1 To UBound(wレコード)
        wsKanjo1 = "": wsGinko1 = ""
        wsKanjo2 = "": wsGinko2 = ""
        For j = 1 To UBound(wレコード(k).xName)
            wsName = wレコード(k).xName(j)
            wsValue = CStr(wレコード(k).xValue(j))
            
            If wsName = "勘定科目" Then
                wsKanjo1 = wsValue
            ElseIf wsName = "銀行番号" Then
                wsGinko1 = wsValue
            End If
        Next j
        
        For lp1 = k + 1 To UBound(wレコード)
            For lp2 = 1 To UBound(wレコード(lp1).xName)
                wsName = wレコード(lp1).xName(lp2)
                wsValue = CStr(wレコード(lp1).xValue(lp2))
                
                If wsName = "勘定科目" Then
                    wsKanjo2 = wsValue
                ElseIf wsName = "銀行番号" Then
                    wsGinko2 = wsValue
                End If
            Next lp2
            
            If wsKanjo1 = wsKanjo2 And wsGinko1 = wsGinko2 Then
                wsMsg = CStr(k) & "行目 勘定科目:" & wsKanjo1 & " 銀行番号:" & wsGinko1 & " と" & vbCr & vbLf
                wsMsg = wsMsg & CStr(lp2) & "行目 勘定科目:" & wsKanjo2 & " 銀行番号:" & wsGinko2 & " を確認してください。"
                MsgBox wsMsg, vbInformation
            End If
        
        Next lp1
    
    Next k
'
    ' =========================================
    '           補助科目マスタUPDATE
    ' =========================================
    wstr = "Delete * from DABA020_補助科目マスタ"
    GDb.Execute wstr
'
    DoEvents
'
    wstr = ""
    wstr = wstr + "Select * From DABA020_補助科目マスタ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)

    For k = 1 To UBound(wレコード)
        wRs.AddNew
        For j = 1 To UBound(wレコード(k).xName)
            wsName = wレコード(k).xName(j)
            wsValue = CStr(wレコード(k).xValue(j))
            
            If wsName = "勘定科目" Then
                wRs("勘定科目") = wsValue
            ElseIf wsName = "勘定科目名" Then
                wRs("勘定科目名") = wsValue
            ElseIf wsName = "銀行番号" Then
                wRs("銀行番号") = wsValue
            ElseIf wsName = "補助科目" Then
                wRs("補助科目") = wsValue
            ElseIf wsName = "補助科目名" Then
                wRs("補助科目名") = wsValue
            End If
        Next j
        wRs.Update
    Next k

    wRs.Close
    Set wRs = Nothing

'
    On Error GoTo 0
'
    MXA040_補助科目取込 = True
'
Exit Function
'----------< ERROR >----------------------------------------------------------------
MXA040_補助科目取込_ERR_CHECK:
    wsMsg = CStr(k) & "行目 " & wsName & ":" & wsValue & " を確認してください。"
    MsgBox wsMsg, vbInformation
    
    Exit Function
'
MXA040_補助科目取込_ERR:
    Err.Clear
    Exit Function

End Function

'------------------------------------------------
' MXA040_補助科目項目名_SET
'------------------------------------------------
Private Sub MXA040_補助科目項目名_SET(pCsvName As String)
'
    '----------< 読込 >-------------------------------------------------------------
    Call MXA040_CsvInit
    '
    Call MXA040_CsvAdd("勘定科目")
    Call MXA040_CsvAdd("勘定科目名")
    Call MXA040_CsvAdd("銀行番号")
    Call MXA040_CsvAdd("銀行名")
    Call MXA040_CsvAdd("補助科目")
    Call MXA040_CsvAdd("補助科目名")
'
    wレコード = MXA040_CsvRead(pCsvName, 2) '1行目タイトル
'
End Sub

'------------------------------------------------
' MXA040_祝日データ項目名_SET
'------------------------------------------------
Private Sub MXA040_祝日データ項目名_SET(pCsvName As String)
'
    '----------< 読込 >-------------------------------------------------------------
    Call MXA040_CsvInit
    '
    Call MXA040_CsvAdd("年月日")
    Call MXA040_CsvAdd("名称")
'
    wレコード = MXA040_CsvRead(pCsvName, 2) '1行目タイトル
'
End Sub

'------------------------------------------------
' MXA040_祝日データ取込
'------------------------------------------------
Public Function MXA040_祝日データ取込(pCsvName As String, pNen As Integer) As Boolean
'
    Dim j As Long, k As Long
    Dim ws01 As String
    Dim wsName As String, wsValue As String
    Dim wsMsg As String
'
    MXA040_祝日データ取込 = False
'
    On Error GoTo MXA040_祝日データ取込_ERR
'
    '項目名で取り込むので名称セット
    Call MXA040_祝日データ項目名_SET(pCsvName)
'
    If UBound(wレコード) = 0 Then
        Exit Function
    End If
'
    For k = 1 To UBound(wレコード)
        For j = 1 To UBound(wレコード(k).xName)
            wsName = wレコード(k).xName(j)
            wsValue = CStr(wレコード(k).xValue(j))
            
            'Data Check
            If wsName = "年月日" Then
                ws01 = Format(wsValue, "yyyy/mm/dd")
                If Not IsDate(ws01) Then
                    GoTo MXA040_祝日データ取込_ERR_CHECK
                Else
                    If Left(wsValue, 4) <> P8.FCStr(pNen) Then
                        GoTo MXA040_祝日データ取込_ERR_CHECK
                    End If
                End If
            End If
        
        Next j
    Next k
'
    ' =========================================
    '         祝日マスタ
    ' =========================================
    wstr = ""
    wstr = wstr & "Delete * "
    wstr = wstr & " From DACA010_祝日マスタ"
    wstr = wstr & " WHERE Format(年月日,'yyyymmdd') >= '" & P8.FCStr(pNen) & "0101" & "'"
    wstr = wstr & " and Format(年月日,'yyyymmdd') <= '" & P8.FCStr(pNen) & "1231" & "'"
    GDb.Execute wstr
'
    wstr = ""
    wstr = wstr & "Select *"
    wstr = wstr & " From DACA010_祝日マスタ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    For k = 1 To UBound(wレコード)
        wRs.AddNew

            For j = 1 To UBound(wレコード(k).xName)
                wsName = wレコード(k).xName(j)
                wsValue = CStr(wレコード(k).xValue(j))
                
                If wsName = "年月日" Then
                    ws01 = Format(wsValue, "yyyy/mm/dd")
                    wRs("年月日") = CDate(ws01)
                ElseIf wsName = "名称" Then
                    wRs("名称") = wsValue
                End If
            
            Next j
    
            wRs("区分") = 0
    
        wRs.Update
    Next k
    
    wRs.Close
    Set wRs = Nothing
'
    On Error GoTo 0
'
    MXA040_祝日データ取込 = True
'
Exit Function
'----------< ERROR >----------------------------------------------------------------
MXA040_祝日データ取込_ERR_CHECK:
    wsMsg = CStr(k) & "行目 " & wsName & ":" & wsValue & " を確認してください。"
    MsgBox wsMsg, vbInformation
    
    Exit Function
'
MXA040_祝日データ取込_ERR:
    Err.Clear
    Exit Function
End Function

'------------------------------------------------
' MXA040_祝日マスタログ一覧
'------------------------------------------------
Public Function MX040_祝日マスタログ一覧(pCsvFileName As String) As Boolean
'
    Dim ws01 As String
    Dim I As Integer
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Function
    End If
    
    DoEvents
'
    If UBound(GCal) <= 0 Then
'        MsgBox "出力するデータがありません", vbInformation
        
        Exit Function
    End If
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        '名称
        Write #1, _
            "借入番号", "銀行番号", "銀行名", "確認年月日"
            
        For I = 1 To UBound(GCal)
            Write #1, _
                GCal(I).借入番号, _
                GCal(I).銀行番号, _
                GCal(I).銀行名, _
                GCal(I).確認年月日
        Next
        
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Function
    End If
'
    MsgBox GRpt.帳票名 & "CSVファイルを出力しました", vbInformation
'
    On Error GoTo 0
'
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Function
    End If
'
End Function

'------------------------------------------------
' MX040_祝日マスタ
'------------------------------------------------
Public Sub MX040_祝日マスタ(pCsvFileName As String, pNendo As String)
'
    Dim ws01 As String
    Dim wDate1 As Date, wDate2 As Date
'
    On Error Resume Next
'
    'schema.ini
    ws01 = wCsvDir & "\" & "schema.ini"
    If Dir(ws01) <> "" Then
        Kill ws01
    End If
'
    'CsvFile Check
    GRet = MX040_CsvFile_Check(pCsvFileName)
    If GRet <> True Then
        FLG_Check = True
        Exit Sub
    End If
    
    DoEvents
'
    wDate1 = CDate(pNendo & "/01/01")
    wDate2 = DateAdd("yyyy", 1, wDate1)
'
    wstr = ""
    wstr = wstr & "Select "
    wstr = wstr & "Format(年月日,'yyyy/mm/dd') As 年月日,"
    wstr = wstr & "名称"
    wstr = wstr & " From DACA010_祝日マスタ"
    wstr = wstr & " Where Format(年月日,'yyyy/mm/dd') >= '" & Format(wDate1, "yyyy/mm/dd") & "'"
    wstr = wstr & " AND Format(年月日,'yyyy/mm/dd') < '" & Format(wDate2, "yyyy/mm/dd") & "'"
    wstr = wstr + " Order By 年月日"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    If wRs.eof Then
        MsgBox "出力するデータがありません", vbInformation
        
        wRs.Close
        Set wRs = Nothing
        
        Exit Sub
    End If
    
    Open wCsvDir & "\" & pCsvFileName For Output Access Write As #1 '出力ファイルを開く
    
    On Error GoTo Err_Hundle
        '名称
        Write #1, _
            "年月日", "名称"
            
    Do Until wRs.eof
        Write #1, _
            P8.FCStr(wRs.Fields("年月日").Value), _
            P8.FCStr(wRs.Fields("名称").Value)
            
    wRs.MoveNext
    Loop
    wRs.Close
    Set wRs = Nothing
    
    Close #1 '出力ファイルを閉じる
'
    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
    MsgBox "出力しました", vbInformation
'
    On Error GoTo 0
'
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
Err_Hundle:
    Close #1 '出力ファイルを閉じる

    If Err.Number Then
        Err.Clear
        FLG_Check = True
        Exit Sub
    End If
'
End Sub
