Attribute VB_Name = "MDA010_仕訳"
Option Explicit
'
Private Const pPROGRAM_ID As String = "MDA010_仕訳"
'
Dim p勘定科目() As MDA010_勘定科目
'------------------------< 勘定科目マスタ >-------------------------
Type MDA010_勘定科目
    仕訳区分 As String
    仕訳補助 As String
    仕訳補助備考 As String
    仕訳名 As String
    社債フラグ As Integer
    
    '振替番号 振替仕訳で使用
    
    借方勘定科目 As String
    借方勘定科目名 As String
    借方銀行番号 As String
    借方補助科目使用 As Integer
    借方個別補助使用 As Integer
    借方補助科目 As String
    借方補助科目名 As String
    貸方勘定科目 As String
    貸方勘定科目名 As String
    貸方銀行番号 As String
    貸方補助科目使用 As Integer
    貸方個別補助使用 As Integer
    貸方補助科目 As String
    貸方補助科目名 As String
    
    '2014/07/31
    伝票番号 As String
    摘要 As String
End Type

Dim p補助科目() As MDA010_補助科目
'------------------------< 勘定科目マスタ >-------------------------
Type MDA010_補助科目
    勘定科目 As String
    勘定科目名 As String
    銀行番号 As String
    補助科目 As String
    補助科目名 As String
End Type

Dim p個別補助() As MDA010_個別補助
'------------------------< 個別勘定科目マスタ >-------------------------
Type MDA010_個別補助
    勘定科目 As String
    勘定科目名 As String
    借入番号 As String
    銀行番号 As String
    個別補助 As String
    個別補助名 As String
End Type

Dim p仕訳データ As MDA010_仕訳データ
Type MDA010_仕訳データ
    借入番号 As String
    銀行番号 As String
    社債フラグ As Integer
    長短区分 As String
    利息区分 As String
    仕訳区分 As String
    借方勘定科目 As String
    貸方勘定科目 As String
End Type

Dim wSDate As Date
'
'------------------------------------------------
' MDA010_勘定科目マスタ設定
'------------------------------------------------
Public Sub MDA010_勘定科目マスタ設定()
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    Dim j As Integer
'
    On Error GoTo MDA010_勘定科目マスタ設定_ERR
'
    ReDim p勘定科目(0)
    
    wstr = ""
    wstr = wstr & "SELECT "
    wstr = wstr & "仕訳区分,仕訳補助,仕訳補助備考,仕訳名,社債フラグ,"
    wstr = wstr & "借方勘定科目,借方勘定科目名,借方補助科目使用,借方個別補助科目使用,"
    wstr = wstr & "貸方勘定科目,貸方勘定科目名,貸方補助科目使用,貸方個別補助科目使用"
    wstr = wstr & " FROM DABA010_勘定科目マスタ"
    wstr = wstr & " WHERE 取消フラグ=0"
    wstr = wstr & " Order by 仕訳区分,社債フラグ,仕訳補助"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.RecordCount = 0 Then
            wRs.Close
            Set wRs = Nothing
            
            Exit Sub
        End If
    
        ReDim p勘定科目(wRs.RecordCount - 1)
    wRs.Close
    Set wRs = Nothing
'
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        j = -1
        
        Do Until wRs.EOF
            j = j + 1
                
            p勘定科目(j).仕訳区分 = P8.FCStr(wRs("仕訳区分"))
            p勘定科目(j).仕訳補助 = P8.FCStr(wRs("仕訳補助"))
            p勘定科目(j).仕訳補助備考 = P8.FCStr(wRs("仕訳補助備考"))
            p勘定科目(j).仕訳名 = P8.FCStr(wRs("仕訳名"))
            p勘定科目(j).社債フラグ = P8.FCDbl(wRs("社債フラグ"))
            p勘定科目(j).借方勘定科目 = P8.FCStr(wRs("借方勘定科目"))
            p勘定科目(j).借方勘定科目名 = P8.FCStr(wRs("借方勘定科目名"))
            p勘定科目(j).借方補助科目使用 = P8.FCDbl(wRs("借方補助科目使用"))
            p勘定科目(j).借方個別補助使用 = P8.FCDbl(wRs("借方個別補助科目使用"))
            p勘定科目(j).貸方勘定科目 = P8.FCStr(wRs("貸方勘定科目"))
            p勘定科目(j).貸方勘定科目名 = P8.FCStr(wRs("貸方勘定科目名"))
            p勘定科目(j).貸方補助科目使用 = P8.FCDbl(wRs("貸方補助科目使用"))
            p勘定科目(j).貸方個別補助使用 = P8.FCDbl(wRs("貸方個別補助科目使用"))
             
            wRs.MoveNext
        Loop
    
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA010_勘定科目マスタ設定_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA010_勘定科目マスタ設定() でエラー" + vbCrLf + vbCrLf + _
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
' MDA010_補助科目マスタ設定
'------------------------------------------------
Public Sub MDA010_補助科目マスタ設定()
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    Dim j As Integer
'
    On Error GoTo MDA010_補助科目マスタ設定_ERR
'
    ReDim p補助科目(0)
    
    wstr = ""
    wstr = wstr & "SELECT "
    wstr = wstr & "勘定科目,勘定科目名,"
    wstr = wstr & "銀行番号,補助科目,補助科目名"
    wstr = wstr & " FROM DABA020_補助科目マスタ"
    wstr = wstr & " Order by 勘定科目,銀行番号,補助科目"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs.RecordCount = 0 Then
            wRs.Close
            Set wRs = Nothing
            
            Exit Sub
        End If
    
        ReDim p補助科目(wRs.RecordCount - 1)
    wRs.Close
    Set wRs = Nothing
'
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        j = -1
        
        Do Until wRs.EOF
            j = j + 1
                
            p補助科目(j).勘定科目 = P8.FCStr(wRs("勘定科目"))
            p補助科目(j).勘定科目名 = P8.FCStr(wRs("勘定科目名"))
            p補助科目(j).銀行番号 = P8.FCStr(wRs("銀行番号"))
            p補助科目(j).補助科目 = P8.FCStr(wRs("補助科目"))
            p補助科目(j).補助科目名 = P8.FCStr(wRs("補助科目名"))
             
            wRs.MoveNext
        Loop
    
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA010_補助科目マスタ設定_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA010_補助科目マスタ設定() でエラー" + vbCrLf + vbCrLf + _
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
' MDA010_個別補助マスタ設定
'------------------------------------------------
Public Sub MDA010_個別補助マスタ設定()
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String
    
    Dim wi01 As Integer, wi02 As Integer, j As Integer
'
    On Error GoTo MDA010_個別補助マスタ設定_ERR
'
    ReDim p個別補助(0)
    
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
        wi01 = wRs2.RecordCount
    wRs2.Close
    Set wRs2 = Nothing
'
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
        If wRs.RecordCount = 0 Then
            wRs.Close
            Set wRs = Nothing
            
            Exit Sub
        End If
        
        wi02 = wRs.RecordCount
        If wi01 * wi02 <= 0 Then
            wRs.Close
            Set wRs = Nothing
            
            Exit Sub
        End If
        
        ReDim p個別補助((wi01 * wi02) - 1)
    wRs.Close
    Set wRs = Nothing
'
    Call AdoRecordsetOpen(GDb, wRs2, wstr2)
        j = -1
        Do Until wRs2.EOF
            Call AdoRecordsetOpen(GDb, wRs, wstr)
                Do Until wRs.EOF
                    j = j + 1
                        
                    p個別補助(j).勘定科目 = P8.FCStr(wRs2("科目"))
                    p個別補助(j).勘定科目名 = P8.FCStr(wRs2("科目名"))
                    p個別補助(j).借入番号 = P8.FCStr(wRs("借入番号"))
                    p個別補助(j).銀行番号 = P8.FCStr(wRs("銀行番号"))
                    p個別補助(j).個別補助 = P8.FCStr(wRs("個別補助科目"))
                    p個別補助(j).個別補助名 = P8.FCStr(wRs("個別補助科目名"))
                     
                    wRs.MoveNext
                Loop
            
            wRs.Close
            Set wRs = Nothing
            
        wRs2.MoveNext
        Loop
    wRs2.Close
    Set wRs2 = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA010_個別補助マスタ設定_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA010_個別補助マスタ設定() でエラー" + vbCrLf + vbCrLf + _
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
' MDA010_勘定科目Read
'------------------------------------------------
Public Function MDA010_勘定科目Read() As MDA010_勘定科目
'
    Dim j As Integer
    Dim wIndex_Kanjo As Integer
    Dim w補助科目 As MDA010_補助科目
    Dim w個別補助 As MDA010_個別補助
'
    On Error GoTo MDA010_勘定科目Read_ERR
'
    MDA010_勘定科目Read.仕訳区分 = ""
    MDA010_勘定科目Read.仕訳補助 = ""
    MDA010_勘定科目Read.仕訳補助備考 = ""
    MDA010_勘定科目Read.仕訳名 = ""
    MDA010_勘定科目Read.社債フラグ = 0
    MDA010_勘定科目Read.借方勘定科目 = ""
    MDA010_勘定科目Read.借方勘定科目名 = ""
    MDA010_勘定科目Read.貸方勘定科目 = ""
    MDA010_勘定科目Read.貸方勘定科目名 = ""
    MDA010_勘定科目Read.借方補助科目使用 = 1
    MDA010_勘定科目Read.借方個別補助使用 = 0
    MDA010_勘定科目Read.借方銀行番号 = ""
    MDA010_勘定科目Read.借方補助科目 = ""
    MDA010_勘定科目Read.借方補助科目名 = ""
    MDA010_勘定科目Read.貸方補助科目使用 = 1
    MDA010_勘定科目Read.貸方個別補助使用 = 0
    MDA010_勘定科目Read.貸方銀行番号 = ""
    MDA010_勘定科目Read.貸方補助科目 = ""
    MDA010_勘定科目Read.貸方補助科目名 = ""
'
    '勘定科目セット
    wIndex_Kanjo = -1
    For j = 0 To UBound(p勘定科目)
        If p勘定科目(j).仕訳区分 = p仕訳データ.仕訳区分 Then
            If p仕訳データ.仕訳区分 = "1" And p勘定科目(j).社債フラグ = p仕訳データ.社債フラグ Then
            '借入金、社債の実行
                If p勘定科目(j).仕訳補助備考 = "長短区分" Then
                    If p勘定科目(j).仕訳補助 = p仕訳データ.長短区分 Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                ElseIf p勘定科目(j).仕訳補助備考 = "利息区分" Then
                    If p勘定科目(j).仕訳補助 = p仕訳データ.利息区分 Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                Else
                '通常は長短区分
                    If p勘定科目(j).仕訳補助 = "9" Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                End If
                
            ElseIf p仕訳データ.仕訳区分 = "2" And p勘定科目(j).社債フラグ = p仕訳データ.社債フラグ Then
            '元金額の支払
                If p勘定科目(j).仕訳補助備考 = "長短区分" Then
                    If p勘定科目(j).仕訳補助 = p仕訳データ.長短区分 Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                ElseIf p勘定科目(j).仕訳補助備考 = "利息区分" Then
                    If p勘定科目(j).仕訳補助 = p仕訳データ.利息区分 Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                Else
                '通常は長短区分
                    If p勘定科目(j).仕訳補助 = "9" Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                End If
            
            ElseIf p仕訳データ.仕訳区分 = "3" And p勘定科目(j).社債フラグ = p仕訳データ.社債フラグ Then
            '利息額の支払
                If p勘定科目(j).仕訳補助備考 = "長短区分" Then
                    If p勘定科目(j).仕訳補助 = p仕訳データ.長短区分 Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                ElseIf p勘定科目(j).仕訳補助備考 = "利息区分" Then
                    If p勘定科目(j).仕訳補助 = p仕訳データ.利息区分 Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                Else
                '通常は長短区分
                    If p勘定科目(j).仕訳補助 = "9" Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                End If
            
            ElseIf p仕訳データ.仕訳区分 = "4" And p勘定科目(j).社債フラグ = p仕訳データ.社債フラグ Then
            '利息額の計上
                If p勘定科目(j).仕訳補助備考 = "長短区分" Then
                    If p勘定科目(j).仕訳補助 = p仕訳データ.長短区分 Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                ElseIf p勘定科目(j).仕訳補助備考 = "利息区分" Then
                    If p勘定科目(j).仕訳補助 = p仕訳データ.利息区分 Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                Else
                '通常は利息区分
                    If p勘定科目(j).仕訳補助 = "9" Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                End If
            
            ElseIf p仕訳データ.仕訳区分 = "5" And p勘定科目(j).社債フラグ = p仕訳データ.社債フラグ Then
            '社債の手数料支払
                If p勘定科目(j).仕訳補助備考 = "長短区分" Then
                    If p勘定科目(j).仕訳補助 = p仕訳データ.長短区分 Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                ElseIf p勘定科目(j).仕訳補助備考 = "利息区分" Then
                    If p勘定科目(j).仕訳補助 = p仕訳データ.利息区分 Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                Else
                '通常は長短区分
                    If p勘定科目(j).仕訳補助 = "9" Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                End If
            
            ElseIf p仕訳データ.仕訳区分 = "6" And p勘定科目(j).社債フラグ = p仕訳データ.社債フラグ Then
            '社債の保証料支払
                If p勘定科目(j).仕訳補助備考 = "長短区分" Then
                    If p勘定科目(j).仕訳補助 = p仕訳データ.長短区分 Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                ElseIf p勘定科目(j).仕訳補助備考 = "利息区分" Then
                    If p勘定科目(j).仕訳補助 = p仕訳データ.利息区分 Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                Else
                '通常は長短区分
                    If p勘定科目(j).仕訳補助 = "9" Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                End If
            
            ElseIf p仕訳データ.仕訳区分 = "7" And p勘定科目(j).社債フラグ = p仕訳データ.社債フラグ Then
            '長期借入金長短振替
                If p勘定科目(j).仕訳補助備考 = "長短区分" Then
                    If p勘定科目(j).仕訳補助 = p仕訳データ.長短区分 Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                ElseIf p勘定科目(j).仕訳補助備考 = "利息区分" Then
                    If p勘定科目(j).仕訳補助 = p仕訳データ.利息区分 Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                Else
                '通常は長短区分
                    If p勘定科目(j).仕訳補助 = "9" Then
                        wIndex_Kanjo = j
                            Exit For
                    End If
                End If
            
            End If
        End If
    Next

    If wIndex_Kanjo <> -1 Then
        MDA010_勘定科目Read.仕訳区分 = p勘定科目(wIndex_Kanjo).仕訳区分
        MDA010_勘定科目Read.仕訳補助 = p勘定科目(wIndex_Kanjo).仕訳補助
        MDA010_勘定科目Read.仕訳補助備考 = p勘定科目(wIndex_Kanjo).仕訳補助備考
        MDA010_勘定科目Read.仕訳名 = p勘定科目(wIndex_Kanjo).仕訳名
        MDA010_勘定科目Read.社債フラグ = p勘定科目(wIndex_Kanjo).社債フラグ

        MDA010_勘定科目Read.借方勘定科目 = p勘定科目(wIndex_Kanjo).借方勘定科目
        MDA010_勘定科目Read.借方勘定科目名 = p勘定科目(wIndex_Kanjo).借方勘定科目名
        MDA010_勘定科目Read.借方補助科目使用 = p勘定科目(wIndex_Kanjo).借方補助科目使用
        MDA010_勘定科目Read.借方個別補助使用 = p勘定科目(wIndex_Kanjo).借方個別補助使用
        MDA010_勘定科目Read.貸方勘定科目 = p勘定科目(wIndex_Kanjo).貸方勘定科目
        MDA010_勘定科目Read.貸方勘定科目名 = p勘定科目(wIndex_Kanjo).貸方勘定科目名
        MDA010_勘定科目Read.貸方補助科目使用 = p勘定科目(wIndex_Kanjo).貸方補助科目使用
        MDA010_勘定科目Read.貸方個別補助使用 = p勘定科目(wIndex_Kanjo).貸方個別補助使用

        If MDA010_勘定科目Read.借方個別補助使用 = 1 Then
            w個別補助 = MDA010_個別補助Read(p勘定科目(wIndex_Kanjo).借方勘定科目, p仕訳データ.銀行番号, p仕訳データ.借入番号)
            MDA010_勘定科目Read.借方銀行番号 = w個別補助.銀行番号
            MDA010_勘定科目Read.借方補助科目 = w個別補助.個別補助
            MDA010_勘定科目Read.借方補助科目名 = w個別補助.個別補助名
        ElseIf MDA010_勘定科目Read.借方補助科目使用 = 1 Then
            w補助科目 = MDA010_補助科目Read(p勘定科目(wIndex_Kanjo).借方勘定科目, p仕訳データ.銀行番号)
            MDA010_勘定科目Read.借方銀行番号 = w補助科目.銀行番号
            MDA010_勘定科目Read.借方補助科目 = w補助科目.補助科目
            MDA010_勘定科目Read.借方補助科目名 = w補助科目.補助科目名
        End If
        
        If MDA010_勘定科目Read.貸方個別補助使用 = 1 Then
            w個別補助 = MDA010_個別補助Read(p勘定科目(wIndex_Kanjo).貸方勘定科目, p仕訳データ.銀行番号, p仕訳データ.借入番号)
            MDA010_勘定科目Read.貸方銀行番号 = w個別補助.銀行番号
            MDA010_勘定科目Read.貸方補助科目 = w個別補助.個別補助
            MDA010_勘定科目Read.貸方補助科目名 = w個別補助.個別補助名
        ElseIf MDA010_勘定科目Read.貸方補助科目使用 = 1 Then
            w補助科目 = MDA010_補助科目Read(p勘定科目(wIndex_Kanjo).貸方勘定科目, p仕訳データ.銀行番号)
            MDA010_勘定科目Read.貸方銀行番号 = w補助科目.銀行番号
            MDA010_勘定科目Read.貸方補助科目 = w補助科目.補助科目
            MDA010_勘定科目Read.貸方補助科目名 = w補助科目.補助科目名
        End If
        
        Exit Function
    Else
        Exit Function
    End If
'

'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA010_勘定科目Read_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA010_勘定科目Read() でエラー" + vbCrLf + vbCrLf + _
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
' MDA010_補助科目Read
'------------------------------------------------
Public Function MDA010_補助科目Read(pKamoku As String, pGinko As String) As MDA010_補助科目
'
    Dim j As Integer
'
    On Error GoTo MDA010_補助科目Read_ERR
'
    MDA010_補助科目Read.勘定科目 = ""
    MDA010_補助科目Read.勘定科目名 = ""
    MDA010_補助科目Read.銀行番号 = ""
    MDA010_補助科目Read.補助科目 = ""
    MDA010_補助科目Read.補助科目名 = ""
'
    For j = 0 To UBound(p補助科目)
        If p補助科目(j).勘定科目 = pKamoku Then
            If p補助科目(j).銀行番号 = pGinko Then
                MDA010_補助科目Read.勘定科目 = p補助科目(j).勘定科目
                MDA010_補助科目Read.勘定科目名 = p補助科目(j).勘定科目名
                MDA010_補助科目Read.銀行番号 = p補助科目(j).銀行番号
                MDA010_補助科目Read.補助科目 = p補助科目(j).補助科目
                MDA010_補助科目Read.補助科目名 = p補助科目(j).補助科目名
                    Exit For
            End If
        End If
    Next
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA010_補助科目Read_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA010_補助科目Read() でエラー" + vbCrLf + vbCrLf + _
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
' MDA010_個別補助Read
'------------------------------------------------
Public Function MDA010_個別補助Read(pKamoku As String, pGinko As String, pBango As String) As MDA010_個別補助
'
    Dim j As Integer
'
    On Error GoTo MDA010_個別補助Read_ERR
'
    MDA010_個別補助Read.勘定科目 = ""
    MDA010_個別補助Read.勘定科目名 = ""
    MDA010_個別補助Read.借入番号 = ""
    MDA010_個別補助Read.銀行番号 = ""
    MDA010_個別補助Read.個別補助 = ""
    MDA010_個別補助Read.個別補助名 = ""
'
    For j = 0 To UBound(p個別補助)
        If p個別補助(j).勘定科目 = pKamoku Then
            If p個別補助(j).銀行番号 = pGinko And p個別補助(j).借入番号 = pBango Then
                MDA010_個別補助Read.勘定科目 = p個別補助(j).勘定科目
                MDA010_個別補助Read.勘定科目名 = p個別補助(j).勘定科目名
                MDA010_個別補助Read.借入番号 = p個別補助(j).借入番号
                MDA010_個別補助Read.銀行番号 = p個別補助(j).銀行番号
                MDA010_個別補助Read.個別補助 = p個別補助(j).個別補助
                MDA010_個別補助Read.個別補助名 = p個別補助(j).個別補助名
                    Exit For
            End If
        End If
    Next
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA010_個別補助Read_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA010_個別補助Read() でエラー" + vbCrLf + vbCrLf + _
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
' MDA010_仕訳現金科目作成
'------------------------------------------------
Public Sub MDA010_仕訳現金科目作成(p借入計画マスタ As MAA910_借入金)
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String

    Dim w勘定科目 As MDA010_勘定科目
    Dim j As Integer
    Dim wDate1 As Date, wDate2 As Date
    Dim p借入金種別 As MAA070_借入金種別
'
    On Error GoTo MDA010_仕訳現金科目作成_ERR
'
    wDate1 = C年月日.平成To西暦("年月日", GRpt.テキスト_01)
    wDate1 = MBA010_締日年月日(wDate1)
    wDate1 = DateAdd("m", -1, wDate1)
    wDate1 = MBA010_締日年月日(Format(wDate1, "yyyy/mm/01"))
    
    wDate2 = C年月日.平成To西暦("年月日", GRpt.テキスト_02)
    wDate2 = MBA010_締日年月日(wDate2)
'
    p仕訳データ.借入番号 = ""
    p仕訳データ.銀行番号 = ""
    p仕訳データ.社債フラグ = 0
    p仕訳データ.長短区分 = ""
    p仕訳データ.利息区分 = ""
    p仕訳データ.仕訳区分 = ""
    p仕訳データ.借方勘定科目 = ""
    p仕訳データ.貸方勘定科目 = ""
'
    '借入金種別区分
    p借入金種別 = MAA070_借入金種別Read(p借入計画マスタ.借入金種別区分)
'
    wstr = ""
    'wstr = wstr & "Select * From DCDA040_仕訳データ2"
    wstr = wstr & "Select * From DCDA040_仕訳データ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        p仕訳データ.社債フラグ = p借入金種別.社債フラグ
        p仕訳データ.借入番号 = p借入計画マスタ.借入番号
        p仕訳データ.長短区分 = p借入計画マスタ.長短区分
        p仕訳データ.利息区分 = p借入計画マスタ.利息区分
        p仕訳データ.銀行番号 = p借入計画マスタ.銀行番号
        
        '借入金の実行　現金科目/借入金
        'If wDate1 <= p借入計画マスタ.実行日 And wDate2 > p借入計画マスタ.実行日 Then
        If wDate1 < p借入計画マスタ.実行日 And wDate2 >= p借入計画マスタ.実行日 Then
            '----------< MDA010_勘定科目Read >----------
            p仕訳データ.仕訳区分 = "1"
            w勘定科目 = MDA010_勘定科目Read()
                
            If w勘定科目.借方勘定科目 <> "" Then
                wRs.AddNew
                    wRs("番号") = 0
                    wRs("年月日") = p借入計画マスタ.実行日
                    wSDate = MBA010_対象年月(CDate(p借入計画マスタ.実行日))
                    wRs("対象年月") = wSDate
                    '
                    wRs("借入番号") = p仕訳データ.借入番号
                    wRs("仕訳区分") = w勘定科目.仕訳区分
                    wRs("仕訳補助") = w勘定科目.仕訳補助
                    wRs("仕訳名") = w勘定科目.仕訳名
                    wRs("社債フラグ") = p仕訳データ.社債フラグ
                    wRs("借方勘定科目") = w勘定科目.借方勘定科目
                    wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                    wRs("借方金額") = p借入計画マスタ.融資金額
                    wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                    wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                    wRs("貸方金額") = p借入計画マスタ.融資金額
                
                    wRs("銀行番号") = w勘定科目.借方銀行番号
                    wRs("借方補助科目") = w勘定科目.借方補助科目
                    wRs("借方補助科目名") = w勘定科目.借方補助科目名
                    wRs("貸方補助科目") = w勘定科目.貸方補助科目
                    wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                
                    If wRs("銀行番号") = "" Then
                        wRs("銀行番号") = p仕訳データ.銀行番号
                    End If
                
                wRs.Update
            End If
        End If

        '借入金の返済　借入金/現金科目
        For j = 1 To UBound(G借入金テーブル)
            
            p仕訳データ.仕訳区分 = ""
            p仕訳データ.借方勘定科目 = ""
            p仕訳データ.貸方勘定科目 = ""
                    
            'If wDate1 <= G借入金テーブル(j).実際年月日 And wDate2 > G借入金テーブル(j).実際年月日 Then
            If wDate1 < G借入金テーブル(j).実際年月日 And wDate2 >= G借入金テーブル(j).実際年月日 Then
                If G借入金テーブル(j).元金額 <> 0 Or G借入金テーブル(j).利息額 <> 0 _
                Or (G借入金テーブル(j).融資残高 <> 0 And p借入計画マスタ.実行日 = G借入金テーブル(j).実際年月日) _
                Or G借入金テーブル(j).保証料 <> 0 Or G借入金テーブル(j).手数料 <> 0 _
                Or Format(p借入計画マスタ.解約実行日, "yyyymmdd") = Format(G借入金テーブル(j).実際年月日, "yyyymmdd") Then '10/06/16 V195
                   
                    '元金額
                    If Format(p借入計画マスタ.解約実行日, "yyyymmdd") = Format(G借入金テーブル(j).実際年月日, "yyyymmdd") _
                    And G借入金テーブル(j).融資残高 <> 0 Then
                    '解約算出
                        '----------< MDA010_勘定科目Read >----------
                        p仕訳データ.仕訳区分 = "2"
                        w勘定科目 = MDA010_勘定科目Read()
                        
                        If w勘定科目.借方勘定科目 <> "" Then
                            wRs.AddNew
                                
                                wRs("番号") = 1
                                wRs("年月日") = G借入金テーブル(j).実際年月日
                                wSDate = MBA010_対象年月(CDate(G借入金テーブル(j).実際年月日))
                                wRs("対象年月") = wSDate
                                '
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ
                                wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("借方金額") = G借入金テーブル(j).融資残高
                                wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("貸方金額") = G借入金テーブル(j).融資残高
                            
                                wRs("借方補助科目") = w勘定科目.借方補助科目
                                wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                wRs("銀行番号") = w勘定科目.貸方銀行番号
                                wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                            
                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If
                            
                            wRs.Update
                        End If
                    ElseIf G借入金テーブル(j).元金額 <> 0 Then
                        '----------< MDA010_勘定科目Read >----------
                        p仕訳データ.仕訳区分 = "2"
                        w勘定科目 = MDA010_勘定科目Read()
                        
                        If w勘定科目.借方勘定科目 <> "" Then
                            wRs.AddNew
                                wRs("番号") = 1
                                wRs("年月日") = G借入金テーブル(j).実際年月日
                                wSDate = MBA010_対象年月(CDate(G借入金テーブル(j).実際年月日))
                                wRs("対象年月") = wSDate
                                '
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ
                                wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("借方金額") = G借入金テーブル(j).元金額
                                wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("貸方金額") = G借入金テーブル(j).元金額
                            
                                wRs("借方補助科目") = w勘定科目.借方補助科目
                                wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                wRs("銀行番号") = w勘定科目.貸方銀行番号
                                wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                            
                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If
                            
                            wRs.Update
                        End If
                    End If
                    
                    If G借入金テーブル(j).利息額 <> 0 Then
                    '利息額　利息前払未払費用/普通預金
                        '----------< MDA010_勘定科目Read >----------
                        p仕訳データ.仕訳区分 = "3"
                        w勘定科目 = MDA010_勘定科目Read()
                        
                        If w勘定科目.借方勘定科目 <> "" Then
                            wRs.AddNew
                                wRs("番号") = 3
                                wRs("年月日") = G借入金テーブル(j).実際年月日
                                wSDate = MBA010_対象年月(CDate(G借入金テーブル(j).実際年月日))
                                wRs("対象年月") = wSDate
                                '
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ
                                wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("借方金額") = G借入金テーブル(j).利息額
                                wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("貸方金額") = G借入金テーブル(j).利息額
                            
                                wRs("借方補助科目") = w勘定科目.借方補助科目
                                wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                wRs("銀行番号") = w勘定科目.貸方銀行番号
                                wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                            
                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If
                                    
                            wRs.Update
                        End If
                    End If
                    
                End If
            End If
                         
            'If wDate2 <= G借入金テーブル(j).実際年月日 Then
            If wDate2 < G借入金テーブル(j).実際年月日 Then
                Exit For
            End If
        Next
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA010_仕訳現金科目作成_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA010_仕訳現金科目作成() でエラー" + vbCrLf + vbCrLf + _
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
' MDA010_仕訳現金科目作成_明細TR
'------------------------------------------------
Public Sub MDA010_仕訳現金科目作成_明細TR(pTbl As String)
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String

    Dim w勘定科目 As MDA010_勘定科目
    Dim wDate1 As Date, wDate2 As Date
    Dim w借入番号 As String
'
    On Error GoTo MDA010_仕訳現金科目作成_明細TR_ERR
'
    wDate1 = C年月日.平成To西暦("年月日", GRpt.テキスト_01)
    wDate1 = MBA010_締日年月日(wDate1)
    wDate1 = DateAdd("m", -1, wDate1)
    wDate1 = MBA010_締日年月日(Format(wDate1, "yyyy/mm/01"))
    
    wDate2 = C年月日.平成To西暦("年月日", GRpt.テキスト_02)
    wDate2 = MBA010_締日年月日(wDate2)
'
    p仕訳データ.社債フラグ = 0
    p仕訳データ.借入番号 = ""
    p仕訳データ.銀行番号 = ""
    p仕訳データ.長短区分 = ""
    p仕訳データ.利息区分 = ""
    p仕訳データ.仕訳区分 = ""
    p仕訳データ.借方勘定科目 = ""
    p仕訳データ.貸方勘定科目 = ""
'
    wstr = ""
    'wstr = wstr + "Select * From DCDA040_仕訳データ2"
    wstr = wstr + "Select * From DCDA040_仕訳データ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        '借入金の実行　現金科目/借入金
        
        w借入番号 = ""
        
        wstr2 = "SELECT "
        wstr2 = wstr2 & "K.実行日,"
        wstr2 = wstr2 & "K.借入番号,"
        wstr2 = wstr2 & "K.銀行番号,"
        wstr2 = wstr2 & "K.長短区分,"
        wstr2 = wstr2 & "K.利息区分,"
        wstr2 = wstr2 & "K.融資金額,"
        wstr2 = wstr2 & "K.解約実行日,"
        wstr2 = wstr2 & "TR.借入番号,"
        wstr2 = wstr2 & "TR.返済回数,"
        wstr2 = wstr2 & "TR.実際年月日,"
        wstr2 = wstr2 & "TR.返済金額,"
        wstr2 = wstr2 & "TR.元金額,"
        wstr2 = wstr2 & "TR.利息額,"
        wstr2 = wstr2 & "TR.融資残高,"
        wstr2 = wstr2 & "TR.保証料,"
        wstr2 = wstr2 & "TR.手数料,"
        wstr2 = wstr2 & "S.社債フラグ"
'        wstr2 = wstr2 & " From (DBDA010_借入金 As K"
        wstr2 = wstr2 & " From (DCIA010_借入金ワーク As K" '16/03/26 利子補給に伴う変更
        wstr2 = wstr2 & " INNER JOIN DBDA010_借入金明細TR As TR"
        wstr2 = wstr2 & " ON K.借入番号 = TR.借入番号)"
        wstr2 = wstr2 & " INNER JOIN DAAA116_借入金種別 As S"
        wstr2 = wstr2 & " ON K.借入金種別区分 = S.借入金種別区分"
        wstr2 = wstr2 & " WHERE K.手入力区分=1 AND K.取消フラグ=0"
        wstr2 = wstr2 & " AND TR.取消フラグ=0 AND TR.取消フラグ２=0"
        wstr2 = wstr2 & " Order by TR.借入番号,TR.返済回数"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
            Do Until wRs2.EOF
                If w借入番号 <> wRs2("K.借入番号") Then
                    p仕訳データ.社債フラグ = P8.FCDbl(wRs2("社債フラグ"))
                    p仕訳データ.借入番号 = wRs2("K.借入番号")
                    p仕訳データ.長短区分 = wRs2("長短区分")
                    p仕訳データ.利息区分 = wRs2("利息区分")
                    p仕訳データ.銀行番号 = wRs2("銀行番号")
                        
                    'If wDate1 <= wRs2("実行日") And wDate2 > wRs2("実行日") Then
                    If wDate1 < wRs2("実行日") And wDate2 >= wRs2("実行日") Then
                        '----------< MDA010_勘定科目Read >----------
                        p仕訳データ.仕訳区分 = "1"
                        w勘定科目 = MDA010_勘定科目Read()
                        
                        If w勘定科目.借方勘定科目 <> "" Then
                            wRs.AddNew
                                wRs("番号") = 0
                                wRs("年月日") = wRs2("実行日")
                                wSDate = MBA010_対象年月(wRs2("実行日"))
                                wRs("対象年月") = wSDate
                                '
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ
                                wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("借方金額") = wRs2("融資金額")
                                wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("貸方金額") = wRs2("融資金額")
                            
                                wRs("銀行番号") = w勘定科目.借方銀行番号
                                wRs("借方補助科目") = w勘定科目.借方補助科目
                                wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                            
                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If
                                    
                            wRs.Update
                        End If
                    End If
                    
                End If
                
                w借入番号 = wRs2("K.借入番号")
                
                wRs2.MoveNext
            Loop
        wRs2.Close
        Set wRs2 = Nothing
        '
        '借入金の返済　借入金/現金科目
        wstr2 = "SELECT "
        wstr2 = wstr2 & "K.実行日,"
        wstr2 = wstr2 & "K.借入番号,"
        wstr2 = wstr2 & "K.銀行番号,"
        wstr2 = wstr2 & "K.長短区分,"
        wstr2 = wstr2 & "K.利息区分,"
        wstr2 = wstr2 & "K.融資金額,"
        wstr2 = wstr2 & "K.解約実行日,"
        wstr2 = wstr2 & "TR.借入番号,"
        wstr2 = wstr2 & "TR.返済回数,"
        wstr2 = wstr2 & "TR.実際年月日,"
        wstr2 = wstr2 & "TR.返済金額,"
        wstr2 = wstr2 & "TR.元金額,"
        wstr2 = wstr2 & "TR.利息額,"
        wstr2 = wstr2 & "TR.融資残高,"
        wstr2 = wstr2 & "TR.保証料,"
        wstr2 = wstr2 & "TR.手数料,"
        wstr2 = wstr2 & "S.社債フラグ"
'        wstr2 = wstr2 & " From (DBDA010_借入金 As K"
        wstr2 = wstr2 & " From (DCIA010_借入金ワーク As K" '16/03/26 利子補給に伴う変更
        wstr2 = wstr2 & " INNER JOIN DBDA010_借入金明細TR As TR"
        wstr2 = wstr2 & " ON K.借入番号 = TR.借入番号)"
        wstr2 = wstr2 & " INNER JOIN DAAA116_借入金種別 As S"
        wstr2 = wstr2 & " ON K.借入金種別区分 = S.借入金種別区分"
        wstr2 = wstr2 & " WHERE K.手入力区分=1 AND K.取消フラグ=0"
        wstr2 = wstr2 & " AND TR.取消フラグ=0 AND TR.取消フラグ２=0"
        wstr2 = wstr2 & " Order by TR.借入番号,TR.返済回数"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
            Do Until wRs2.EOF
                
                p仕訳データ.社債フラグ = P8.FCDbl(wRs2("社債フラグ"))
                p仕訳データ.借入番号 = wRs2("K.借入番号")
                p仕訳データ.長短区分 = wRs2("長短区分")
                p仕訳データ.利息区分 = wRs2("利息区分")
                p仕訳データ.銀行番号 = wRs2("銀行番号")
                
                'If wDate1 <= wRs2("実際年月日") And wDate2 > wRs2("実際年月日") Then
                If wDate1 < wRs2("実際年月日") And wDate2 >= wRs2("実際年月日") Then
                    If wRs2("元金額") <> 0 Or wRs2("利息額") <> 0 _
                    Or wRs2("融資残高") <> 0 And wRs2("実行日") = wRs2("実際年月日") _
                    Or wRs2("保証料") <> 0 Or wRs2("手数料") <> 0 _
                    Or Format(wRs2("解約実行日"), "yyyymmdd") = Format(wRs2("実際年月日"), "yyyymmdd") Then '10/06/16 V195
                        
                        '元金額
                        If Format(wRs2("解約実行日"), "yyyymmdd") = Format(wRs2("実際年月日"), "yyyymmdd") _
                        And wRs2("融資残高") <> 0 Then
                        '解約算出
                            '----------< MDA010_勘定科目Read >----------
                            p仕訳データ.仕訳区分 = "2"
                            w勘定科目 = MDA010_勘定科目Read()
                            
                            If w勘定科目.借方勘定科目 <> "" Then
                                wRs.AddNew
                                    wRs("番号") = 1
                                    wRs("年月日") = wRs2("実際年月日")
                                    wSDate = MBA010_対象年月(wRs2("実際年月日"))
                                    wRs("対象年月") = wSDate
                                    '
                                    wRs("借入番号") = p仕訳データ.借入番号
                                    wRs("仕訳区分") = w勘定科目.仕訳区分
                                    wRs("仕訳補助") = w勘定科目.仕訳補助
                                    wRs("仕訳名") = w勘定科目.仕訳名
                                    wRs("社債フラグ") = p仕訳データ.社債フラグ
                                    wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                    wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                    wRs("借方金額") = wRs2("融資残高")
                                    wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                    wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                    wRs("貸方金額") = wRs2("融資残高")
                                
                                    wRs("借方補助科目") = w勘定科目.借方補助科目
                                    wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                    wRs("銀行番号") = w勘定科目.貸方銀行番号
                                    wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                    wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                                
                                    If wRs("銀行番号") = "" Then
                                        wRs("銀行番号") = p仕訳データ.銀行番号
                                    End If
                                            
                                wRs.Update
                            End If
                        
                        ElseIf wRs2("元金額") <> 0 Then
                            '----------< MDA010_勘定科目Read >----------
                            p仕訳データ.仕訳区分 = "2"
                            w勘定科目 = MDA010_勘定科目Read()
                            
                            If w勘定科目.借方勘定科目 <> "" Then
                                wRs.AddNew
                                    wRs("番号") = 1
                                    wRs("年月日") = wRs2("実際年月日")
                                    wSDate = MBA010_対象年月(wRs2("実際年月日"))
                                    wRs("対象年月") = wSDate
                                    '
                                    wRs("借入番号") = p仕訳データ.借入番号
                                    wRs("仕訳区分") = w勘定科目.仕訳区分
                                    wRs("仕訳補助") = w勘定科目.仕訳補助
                                    wRs("仕訳名") = w勘定科目.仕訳名
                                    wRs("社債フラグ") = p仕訳データ.社債フラグ
                                    wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                    wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                    wRs("借方金額") = wRs2("元金額")
                                    wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                    wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                    wRs("貸方金額") = wRs2("元金額")
                                
                                    wRs("借方補助科目") = w勘定科目.借方補助科目
                                    wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                    wRs("銀行番号") = w勘定科目.貸方銀行番号
                                    wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                    wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                                
                                    If wRs("銀行番号") = "" Then
                                        wRs("銀行番号") = p仕訳データ.銀行番号
                                    End If
                                            
                                wRs.Update
                            End If
                        End If
                        
                        If wRs2("利息額") <> 0 Then
                        '利息額　利息前払未払費用/普通預金
                            '----------< MDA010_勘定科目Read >----------
                            p仕訳データ.仕訳区分 = "3"
                            w勘定科目 = MDA010_勘定科目Read()
                            
                            If w勘定科目.借方勘定科目 <> "" Then
                                wRs.AddNew
                                    wRs("番号") = 3
                                    wRs("年月日") = wRs2("実際年月日")
                                    wSDate = MBA010_対象年月(wRs2("実際年月日"))
                                    wRs("対象年月") = wSDate
                                    '
                                    wRs("借入番号") = p仕訳データ.借入番号
                                    wRs("仕訳区分") = w勘定科目.仕訳区分
                                    wRs("仕訳補助") = w勘定科目.仕訳補助
                                    wRs("仕訳名") = w勘定科目.仕訳名
                                    wRs("社債フラグ") = p仕訳データ.社債フラグ
                                    wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                    wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                    wRs("借方金額") = wRs2("利息額")
                                    wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                    wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                    wRs("貸方金額") = wRs2("利息額")
                                
                                    wRs("借方補助科目") = w勘定科目.借方補助科目
                                    wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                    wRs("銀行番号") = w勘定科目.貸方銀行番号
                                    wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                    wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                                
                                    If wRs("銀行番号") = "" Then
                                        wRs("銀行番号") = p仕訳データ.銀行番号
                                    End If
                                            
                                wRs.Update
                            End If
                        End If
                        
                    End If
                  
                End If
                
'                If wDate2 <= wRs2("実際年月日") Then
'                    Exit Do
'                End If
                
                wRs2.MoveNext
            Loop
        wRs2.Close
        Set wRs2 = Nothing
        '
        '借入金の返済　借入金/現金科目
        wstr2 = "SELECT "
        wstr2 = wstr2 & "K.実行日,"
        wstr2 = wstr2 & "K.借入番号,"
        wstr2 = wstr2 & "K.銀行番号,"
        wstr2 = wstr2 & "K.長短区分,"
        wstr2 = wstr2 & "K.利息区分,"
        wstr2 = wstr2 & "TR2.借入番号,"
        wstr2 = wstr2 & "TR2.実際年月日,"
        wstr2 = wstr2 & "TR2.保証料,"
        wstr2 = wstr2 & "TR2.初期手数料+TR2.元金手数料+TR2.利息手数料 As 手数料,"
        wstr2 = wstr2 & "S.社債フラグ"
'        wstr2 = wstr2 & " From (DBDA010_借入金 As K"
        wstr2 = wstr2 & " From (DCIA010_借入金ワーク As K" '16/03/26 利子補給に伴う変更
        wstr2 = wstr2 & " INNER JOIN DBDA010_借入金明細TR2 As TR2"
        wstr2 = wstr2 & " ON K.借入番号 = TR2.借入番号)"
        wstr2 = wstr2 & " INNER JOIN DAAA116_借入金種別 As S"
        wstr2 = wstr2 & " ON K.借入金種別区分 = S.借入金種別区分"
        wstr2 = wstr2 & " WHERE S.社債フラグ=1 AND K.取消フラグ=0"
        wstr2 = wstr2 & " And K.手入力区分=1 AND K.取消フラグ=0"
        wstr2 = wstr2 & " AND TR2.取消フラグ=0 AND TR2.取消フラグ２=0"
        wstr2 = wstr2 & " Order by TR2.借入番号,TR2.実際年月日"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
            Do Until wRs2.EOF
                p仕訳データ.社債フラグ = P8.FCDbl(wRs2("社債フラグ"))
                p仕訳データ.借入番号 = wRs2("K.借入番号")
                p仕訳データ.長短区分 = wRs2("長短区分")
                p仕訳データ.利息区分 = wRs2("利息区分")
                p仕訳データ.銀行番号 = wRs2("銀行番号")
    
                'If wDate1 <= wRs2("実際年月日") And wDate2 > wRs2("実際年月日") Then
                If wDate1 < wRs2("実際年月日") And wDate2 >= wRs2("実際年月日") Then
                    If wRs2("手数料") <> 0 Then
                    '手数料　手数料/普通預金
                        '----------< MDA010_勘定科目Read >----------
                        p仕訳データ.仕訳区分 = "5"
                        w勘定科目 = MDA010_勘定科目Read()
                        
                        If w勘定科目.借方勘定科目 <> "" Then
                            wRs.AddNew
                                wRs("番号") = 3
                                wRs("年月日") = wRs2("実際年月日")
                                wSDate = MBA010_対象年月(wRs2("実際年月日"))
                                wRs("対象年月") = wSDate
                                '
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ
                                wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("借方金額") = wRs2("手数料")
                                wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("貸方金額") = wRs2("手数料")
                            
                                wRs("借方補助科目") = w勘定科目.借方補助科目
                                wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                wRs("銀行番号") = w勘定科目.貸方銀行番号
                                wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                            
                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If
                                    
                            wRs.Update
                        End If
                    End If
                    
                    If wRs2("保証料") <> 0 Then
                    '保証料　保証料/普通預金
                        '----------< MDA010_勘定科目Read >----------
                        p仕訳データ.仕訳区分 = "6"
                        w勘定科目 = MDA010_勘定科目Read()
                        
                        If w勘定科目.借方勘定科目 <> "" Then
                            wRs.AddNew
                                wRs("番号") = 3
                                wRs("年月日") = wRs2("実際年月日")
                                wSDate = MBA010_対象年月(wRs2("実際年月日"))
                                wRs("対象年月") = wSDate
                                '
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ
                                wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("借方金額") = wRs2("保証料")
                                wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("貸方金額") = wRs2("保証料")
                            
                                wRs("借方補助科目") = w勘定科目.借方補助科目
                                wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                wRs("銀行番号") = w勘定科目.貸方銀行番号
                                wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                            
                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If
                                    
                            wRs.Update
                        End If
                    End If
                End If
                    
                wRs2.MoveNext
            Loop
        wRs2.Close
        Set wRs2 = Nothing
        '
    
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA010_仕訳現金科目作成_明細TR_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA010_仕訳現金科目作成_明細TR() でエラー" + vbCrLf + vbCrLf + _
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
' MDA010_仕訳計上科目作成
'------------------------------------------------
Public Sub MDA010_仕訳計上科目作成()
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String

    Dim w勘定科目 As MDA010_勘定科目
    Dim wDate1 As Date, wDate2 As Date
'
    On Error GoTo MDA010_仕訳計上科目作成_ERR
'
    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    wDate1 = C年月日.平成To西暦("年月日", GRpt.テキスト_01)
    wDate2 = C年月日.平成To西暦("年月日", GRpt.テキスト_02)
    wDate2 = DateAdd("m", 1, wDate2)
'
    p仕訳データ.社債フラグ = 0
    p仕訳データ.借入番号 = ""
    p仕訳データ.銀行番号 = ""
    p仕訳データ.長短区分 = ""
    p仕訳データ.利息区分 = ""
    p仕訳データ.仕訳区分 = ""
    p仕訳データ.借方勘定科目 = ""
    p仕訳データ.貸方勘定科目 = ""
'
    wstr = ""
    wstr = wstr + "Select * From DCDA040_仕訳データ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        wstr2 = "SELECT "
        wstr2 = wstr2 & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',IIF(利息額減<>0,1,2), IIF(利息額減<>0,2,1)) As 日番号,"
        wstr2 = wstr2 & "K.借入番号,"
        wstr2 = wstr2 & "K.銀行番号,"
        wstr2 = wstr2 & "K.長短区分,"
        wstr2 = wstr2 & "K.利息区分,"
        wstr2 = wstr2 & "M.返済年月日,"
        wstr2 = wstr2 & "利息額増,"
        wstr2 = wstr2 & "利息額減,"
        wstr2 = wstr2 & "IIF(利息額増<>0,利息額増,利息額減) As 利息額,"
        wstr2 = wstr2 & "S.社債フラグ"
        wstr2 = wstr2 & " From (DCDA030_利息未払前払明細 As M"
        wstr2 = wstr2 & " INNER JOIN DBDA010_借入金 As K"
        wstr2 = wstr2 & " ON K.借入番号 = M.借入番号)"
        wstr2 = wstr2 & " INNER JOIN DAAA116_借入金種別 As S"
        wstr2 = wstr2 & " ON K.借入金種別区分 = S.借入金種別区分"
        wstr2 = wstr2 & " Where Format(締年月,'yyyy/mm')>='" & Format(wDate1, "yyyy/mm") & "'"
        wstr2 = wstr2 & " And Format(締年月,'yyyy/mm')<'" & Format(wDate2, "yyyy/mm") & "'"
        wstr2 = wstr2 & " order by M.借入番号,返済年月日,1"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
            Do Until wRs2.EOF
                p仕訳データ.社債フラグ = P8.FCDbl(wRs2("社債フラグ"))
                p仕訳データ.借入番号 = wRs2("借入番号")
                p仕訳データ.長短区分 = wRs2("長短区分")
                p仕訳データ.利息区分 = wRs2("利息区分")
                p仕訳データ.銀行番号 = wRs2("銀行番号")

                If wRs2("利息額") <> 0 Then
                '利息額　支払利息/利息前払未払費用
                    '----------< MDA010_勘定科目Read >----------
                    p仕訳データ.仕訳区分 = "4"
                    w勘定科目 = MDA010_勘定科目Read()

                    If w勘定科目.借方勘定科目 <> "" Then
                        wRs.AddNew
                        
                            wRs("番号") = 4
                            wRs("年月日") = wRs2("返済年月日")
                            wSDate = MBA010_対象年月(wRs2("返済年月日"))
                            wRs("対象年月") = wSDate
                            '
                            wRs("日番号") = wRs2("日番号")
                            wRs("借入番号") = p仕訳データ.借入番号
                            wRs("仕訳区分") = w勘定科目.仕訳区分
                            wRs("仕訳補助") = w勘定科目.仕訳補助
                            wRs("仕訳名") = w勘定科目.仕訳名
                            wRs("社債フラグ") = p仕訳データ.社債フラグ
                            
                            If p仕訳データ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
                                If wRs2("利息額減") <> 0 Then
                                '計上
                                    wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                    wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                    wRs("借方金額") = wRs2("利息額")
                                    wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                    wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                    wRs("貸方金額") = wRs2("利息額")
            
                                    wRs("銀行番号") = w勘定科目.借方銀行番号
                                    wRs("借方補助科目") = w勘定科目.借方補助科目
                                    wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                    wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                    wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                                Else
                                    wRs("借方勘定科目") = w勘定科目.貸方勘定科目
                                    wRs("借方勘定科目名") = w勘定科目.貸方勘定科目名
                                    wRs("借方金額") = wRs2("利息額")
                                    wRs("貸方勘定科目") = w勘定科目.借方勘定科目
                                    wRs("貸方勘定科目名") = w勘定科目.借方勘定科目名
                                    wRs("貸方金額") = wRs2("利息額")
            
                                    wRs("銀行番号") = w勘定科目.借方銀行番号
                                    wRs("借方補助科目") = w勘定科目.貸方補助科目
                                    wRs("借方補助科目名") = w勘定科目.貸方補助科目名
                                    wRs("貸方補助科目") = w勘定科目.借方補助科目
                                    wRs("貸方補助科目名") = w勘定科目.借方補助科目名
                                End If
                                
                            ElseIf p仕訳データ.利息区分 = XMXA020_区分("利息区分", "利息後払") Then
                                If wRs2("利息額増") <> 0 Then
                                '計上
                                    wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                    wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                    wRs("借方金額") = wRs2("利息額")
                                    wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                    wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                    wRs("貸方金額") = wRs2("利息額")
            
                                    wRs("銀行番号") = w勘定科目.借方銀行番号
                                    wRs("借方補助科目") = w勘定科目.借方補助科目
                                    wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                    wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                    wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                                Else
                                    wRs("借方勘定科目") = w勘定科目.貸方勘定科目
                                    wRs("借方勘定科目名") = w勘定科目.貸方勘定科目名
                                    wRs("借方金額") = wRs2("利息額")
                                    wRs("貸方勘定科目") = w勘定科目.借方勘定科目
                                    wRs("貸方勘定科目名") = w勘定科目.借方勘定科目名
                                    wRs("貸方金額") = wRs2("利息額")
            
                                    wRs("銀行番号") = w勘定科目.借方銀行番号
                                    wRs("借方補助科目") = w勘定科目.貸方補助科目
                                    wRs("借方補助科目名") = w勘定科目.貸方補助科目名
                                    wRs("貸方補助科目") = w勘定科目.借方補助科目
                                    wRs("貸方補助科目名") = w勘定科目.借方補助科目名
                                End If
                            End If
                        
                        wRs.Update
                    End If
                End If

                wRs2.MoveNext
            Loop
        wRs2.Close
        Set wRs2 = Nothing
    
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA010_仕訳計上科目作成_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA010_仕訳計上科目作成() でエラー" + vbCrLf + vbCrLf + _
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
' MDA010_仕訳計上科目作成_残高
'------------------------------------------------
Public Sub MDA010_仕訳計上科目作成_残高()
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String

    Dim w勘定科目 As MDA010_勘定科目
    Dim wDate1 As Date, wDate2 As Date
    Dim w番号 As String
'
    On Error GoTo MDA010_仕訳計上科目作成_残高_ERR
'
    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    wDate1 = C年月日.平成To西暦("年月日", GRpt.テキスト_01)
    wDate1 = MBA010_締日年月日(wDate1)
    wDate2 = DateAdd("d", 1, wDate1)
    
    w番号 = Right("00" + CStr(GInt1), 2)
'
    p仕訳データ.社債フラグ = 0
    p仕訳データ.借入番号 = ""
    p仕訳データ.銀行番号 = ""
    p仕訳データ.長短区分 = ""
    p仕訳データ.利息区分 = ""
    p仕訳データ.仕訳区分 = ""
    p仕訳データ.借方勘定科目 = ""
    p仕訳データ.貸方勘定科目 = ""

    w番号 = Right("00" & CStr(GInt1), 2)
'
    'ワークデータ 計上仕訳作成と振り戻し仕訳作成
    wstr = ""
    wstr = wstr + "Select * From DCDA040_仕訳データ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        wstr2 = ""
        wstr2 = wstr2 & "SELECT Z.借入番号,K.銀行番号, K.利息区分,K.長短区分, S.社債フラグ,"
        wstr2 = wstr2 & " IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息_" & w番号 & ",Z.未払利息_" & w番号 & ") AS 利息残高"
        wstr2 = wstr2 & " FROM (DCDA010_借入残高推移表結果 AS Z"
        wstr2 = wstr2 & " INNER JOIN DBDA010_借入金 AS K"
        wstr2 = wstr2 & " ON Z.借入番号 = K.借入番号)"
        wstr2 = wstr2 & " LEFT JOIN DAAA116_借入金種別 AS S"
        wstr2 = wstr2 & " ON K.借入金種別区分 = S.借入金種別区分"
        wstr2 = wstr2 & " Where IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息_" & w番号 & ",Z.未払利息_" & w番号 & ")<>0"
        wstr2 = wstr2 & " ORDER BY Z.借入番号"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
            Do Until wRs2.EOF
                p仕訳データ.社債フラグ = P8.FCDbl(wRs2("社債フラグ"))
                p仕訳データ.借入番号 = wRs2("借入番号")
                p仕訳データ.長短区分 = wRs2("利息区分")
                p仕訳データ.利息区分 = wRs2("利息区分")
                p仕訳データ.銀行番号 = wRs2("銀行番号")

                '利息額　支払利息/利息前払未払費用
                    '----------< MDA010_勘定科目Read >----------
                    p仕訳データ.仕訳区分 = "4"
                    w勘定科目 = MDA010_勘定科目Read()

                    If w勘定科目.借方勘定科目 <> "" Then
                        If p仕訳データ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
                            wRs.AddNew
                                wRs("番号") = 4
                                wRs("年月日") = wDate1
                                wSDate = MBA010_対象年月(wDate1)
                                wRs("対象年月") = wSDate
                                '
                                wRs("日番号") = 0
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ

                                wRs("借方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("借方金額") = wRs2("利息残高")
                                wRs("貸方勘定科目") = w勘定科目.借方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("貸方金額") = wRs2("利息残高")

                                wRs("銀行番号") = w勘定科目.借方銀行番号
                                wRs("借方補助科目") = w勘定科目.貸方補助科目
                                wRs("借方補助科目名") = w勘定科目.貸方補助科目名
                                wRs("貸方補助科目") = w勘定科目.借方補助科目
                                wRs("貸方補助科目名") = w勘定科目.借方補助科目名

                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If

                            wRs.Update

                            '振り戻し作成
                            wRs.AddNew

                                wRs("番号") = 4
                                wRs("年月日") = wDate2
                                wSDate = MBA010_対象年月(wDate2)
                                wRs("対象年月") = wSDate
                                '
                                wRs("日番号") = 0
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名 & "/振り戻し"
                                wRs("社債フラグ") = p仕訳データ.社債フラグ

                                wRs("銀行番号") = w勘定科目.借方銀行番号
                                wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                                wRs("借方補助科目") = w勘定科目.借方補助科目
                                wRs("借方補助科目名") = w勘定科目.借方補助科目名

                                wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("貸方金額") = wRs2("利息残高")
                                wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("借方金額") = wRs2("利息残高")

                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If

                            wRs.Update

                        ElseIf p仕訳データ.利息区分 = XMXA020_区分("利息区分", "利息後払") Then
                            wRs.AddNew
    
                                wRs("番号") = 4
                                wRs("年月日") = wDate1
                                wSDate = MBA010_対象年月(wDate1)
                                wRs("対象年月") = wSDate
                                '
                                wRs("日番号") = 0
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ
    
                                wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("借方金額") = wRs2("利息残高")
                                wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("貸方金額") = wRs2("利息残高")
    
                                wRs("銀行番号") = w勘定科目.借方銀行番号
                                wRs("借方補助科目") = w勘定科目.借方補助科目
                                wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
    
                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If
    
                            wRs.Update
    
                            '振り戻し作成
                            wRs.AddNew
    
                                wRs("番号") = 4
                                wRs("年月日") = wDate2
                                wSDate = MBA010_対象年月(wDate2)
                                wRs("対象年月") = wSDate
                                '
                                wRs("日番号") = 0
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名 & "/振り戻し"
                                wRs("社債フラグ") = p仕訳データ.社債フラグ
    
                                wRs("銀行番号") = w勘定科目.借方銀行番号
                                wRs("貸方補助科目") = w勘定科目.借方補助科目
                                wRs("貸方補助科目名") = w勘定科目.借方補助科目名
                                wRs("借方補助科目") = w勘定科目.貸方補助科目
                                wRs("借方補助科目名") = w勘定科目.貸方補助科目名
    
                                wRs("貸方勘定科目") = w勘定科目.借方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("貸方金額") = wRs2("利息残高")
                                wRs("借方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("借方金額") = wRs2("利息残高")
    
                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If
    
                            wRs.Update
                        End If
                    End If

                wRs2.MoveNext
            Loop
        wRs2.Close
        Set wRs2 = Nothing
    
    
    
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA010_仕訳計上科目作成_残高_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA010_仕訳計上科目作成_残高() でエラー" + vbCrLf + vbCrLf + _
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
' MDA010_仕訳長短振替科目作成
'------------------------------------------------
Public Sub MDA010_仕訳長短振替科目作成()
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String

    Dim w勘定科目 As MDA010_勘定科目
    Dim wDate1 As Date, wDate2 As Date
    Dim j As Integer, wiCnt As Integer, w間隔 As Integer
    Dim w番号 As String, ws01 As String, ws02 As String
'
    On Error GoTo MDA010_仕訳長短振替科目作成_ERR
'
    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    wDate1 = C年月日.平成To西暦("年月日", GRpt.テキスト_01)
    wDate1 = MBA010_締日年月日(wDate1)
    wDate2 = DateAdd("d", 1, wDate1)
'
    Select Case G基本情報.決算サイクル
    Case 1
    '月次決算
        '仮の決算月を指定付に設定し、G基本情報.決算サイクル=年次と同処理をする
        w間隔 = 12
    Case 3
        w間隔 = G基本情報.決算サイクル
    Case 6
        w間隔 = G基本情報.決算サイクル
    Case Else
        w間隔 = 12
    End Select
    
    wiCnt = 12 / w間隔
    wiCnt = GInt1 + wiCnt - 1
    
    If wiCnt > 12 Then
        wiCnt = 12
    End If
'
    p仕訳データ.社債フラグ = 0
    p仕訳データ.借入番号 = ""
    p仕訳データ.銀行番号 = ""
    p仕訳データ.長短区分 = ""
    p仕訳データ.利息区分 = ""
    p仕訳データ.仕訳区分 = ""
    p仕訳データ.借方勘定科目 = ""
    p仕訳データ.貸方勘定科目 = ""
'
    '銀行 金額集計/振替仕訳
    wstr = ""
    wstr = wstr + "Select * From DCDA040_仕訳データ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        wstr2 = ""
        wstr2 = wstr2 & "SELECT KS.社債フラグ,K.銀行番号,S.借入番号,"
        wstr2 = wstr2 & " K.長短区分,K.利息区分,"
        'wstr2 = wstr2 & " (S.元金_" & w番号 & "+S.元金_" & w番号2 & ") AS 元金額" '1年間（半期+半期)日本ガス
        
            ws01 = ""
            For j = GInt1 To wiCnt - 1
                w番号 = Right("00" + CStr(j), 2)
                ws01 = ws01 & "S.元金_" & w番号 & "+"
            Next j
            
            w番号 = Right("00" + CStr(wiCnt), 2)
            ws01 = ws01 & "S.元金_" & w番号
        
        wstr2 = wstr2 & " (" & ws01 & ") AS 元金額" '1年間
            
        wstr2 = wstr2 & " FROM (DCDA010_借入残高推移表結果 AS S"
        wstr2 = wstr2 & " INNER JOIN DBDA010_借入金 AS K"
        wstr2 = wstr2 & " ON S.借入番号 = K.借入番号)"
        wstr2 = wstr2 & " INNER JOIN DAAA116_借入金種別 KS"
        wstr2 = wstr2 & " ON K.借入金種別区分 = KS.借入金種別区分"
        wstr2 = wstr2 & " Where K.長短区分=" & P8.FCDbl(XMXA020_区分("長短区分", "長期借入金"))
        'wstr2 = wstr2 & " And (S.元金_" & w番号 & " <> 0 Or S.元金_" & w番号2 & " <> 0)"'1年間（半期+半期)日本ガス
        
        '実行日
        wstr2 = wstr2 & " And format(K.実行日,'yyyy/mm/dd')<'" & Format(wDate2, "yyyy/mm/dd") & "'"
            
            ws02 = ""
            For j = GInt1 To wiCnt - 1
                w番号 = Right("00" + CStr(j), 2)
                ws02 = ws02 & "S.元金_" & w番号 & "<> 0 Or "
            Next j
            
            w番号 = Right("00" + CStr(wiCnt), 2)
            ws02 = ws02 & "S.元金_" & w番号 & "<> 0"
        
        wstr2 = wstr2 & " And (" & ws02 & ")" '1年間
        
        wstr2 = wstr2 & " ORDER BY KS.社債フラグ, K.銀行番号,S.借入番号"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
        Do Until wRs2.EOF
            
                p仕訳データ.社債フラグ = P8.FCDbl(wRs2("社債フラグ"))
                p仕訳データ.借入番号 = wRs2("借入番号")
                p仕訳データ.銀行番号 = wRs2("銀行番号")
                p仕訳データ.長短区分 = wRs2("長短区分")
                p仕訳データ.利息区分 = wRs2("利息区分")
                
                '----------< MDA010_勘定科目Read >----------
                p仕訳データ.仕訳区分 = "7"
                w勘定科目 = MDA010_勘定科目Read()
                
                If w勘定科目.借方勘定科目 <> "" Then
                    wRs.AddNew
                        
                        wRs("番号") = 1
                        wRs("年月日") = wDate1
                        wSDate = MBA010_対象年月(wDate1)
                        wRs("対象年月") = wSDate
                        '
                        wRs("日番号") = 0
                        wRs("借入番号") = p仕訳データ.借入番号
                        wRs("仕訳区分") = w勘定科目.仕訳区分
                        wRs("仕訳補助") = w勘定科目.仕訳補助
                        wRs("仕訳名") = w勘定科目.仕訳名
                        wRs("社債フラグ") = p仕訳データ.社債フラグ

                        wRs("借方勘定科目") = w勘定科目.借方勘定科目
                        wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                        wRs("借方金額") = wRs2("元金額")
                        wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                        wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                        wRs("貸方金額") = wRs2("元金額")

                        wRs("借方補助科目") = w勘定科目.借方補助科目
                        wRs("借方補助科目名") = w勘定科目.借方補助科目名
                        wRs("銀行番号") = w勘定科目.貸方銀行番号
                        wRs("貸方補助科目") = w勘定科目.貸方補助科目
                        wRs("貸方補助科目名") = w勘定科目.貸方補助科目名

                        If wRs("銀行番号") = "" Then
                            wRs("銀行番号") = p仕訳データ.銀行番号
                        End If
                    
                    wRs.Update
                
                    '振り戻し作成
                    wRs.AddNew

                        wRs("番号") = 1
                        wRs("年月日") = wDate2
                        wSDate = MBA010_対象年月(wDate2)
                        wRs("対象年月") = wSDate
                        '
                        wRs("日番号") = 0
                        wRs("借入番号") = p仕訳データ.借入番号
                        wRs("仕訳区分") = w勘定科目.仕訳区分
                        wRs("仕訳補助") = w勘定科目.仕訳補助
                        wRs("仕訳名") = w勘定科目.仕訳名 & "/振り戻し"
                        wRs("社債フラグ") = p仕訳データ.社債フラグ

                        wRs("貸方勘定科目") = w勘定科目.借方勘定科目
                        wRs("貸方勘定科目名") = w勘定科目.借方勘定科目名
                        wRs("貸方金額") = wRs2("元金額")
                        wRs("借方勘定科目") = w勘定科目.貸方勘定科目
                        wRs("借方勘定科目名") = w勘定科目.貸方勘定科目名
                        wRs("借方金額") = wRs2("元金額")

                        wRs("貸方補助科目") = w勘定科目.借方補助科目
                        wRs("貸方補助科目名") = w勘定科目.借方補助科目名
                        wRs("銀行番号") = w勘定科目.貸方銀行番号
                        wRs("借方補助科目") = w勘定科目.貸方補助科目
                        wRs("借方補助科目名") = w勘定科目.貸方補助科目名
                        
                        If wRs("銀行番号") = "" Then
                            wRs("銀行番号") = p仕訳データ.銀行番号
                        End If
        
                    wRs.Update
                End If
            
            wRs2.MoveNext
        Loop
        wRs2.Close
        Set wRs2 = Nothing
        
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA010_仕訳長短振替科目作成_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA010_仕訳長短振替科目作成() でエラー" + vbCrLf + vbCrLf + _
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
' MDA010_月次仕訳作成_日本ガス
'------------------------------------------------
Public Sub MDA010_月次仕訳作成_日本ガス()
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String

    Dim w勘定科目 As MDA010_勘定科目
    Dim wDate1 As Date, wDate2 As Date
    Dim w番号 As String
'
    On Error GoTo MDA010_月次仕訳作成_日本ガス_ERR
'
    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    '利息額 以外 以降
    wstr = ""
    wstr = "INSERT INTO DCDA040_仕訳データ"
    wstr = wstr & " Select * From DCDA040_仕訳データ2"
    wstr = wstr & " Where 仕訳区分<>'3'"
    GDb.Execute wstr
    
    '利息額は集計
    wstr = ""
    wstr = wstr + "Select * From DCDA040_仕訳データ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        wstr2 = ""
        wstr2 = wstr2 & "SELECT 番号,年月日,日番号,銀行番号,仕訳区分,仕訳補助,仕訳名,社債フラグ,"
        wstr2 = wstr2 & " 借方勘定科目,借方勘定科目名,借方補助科目,借方補助科目名, Sum(借方金額) AS 借方金額合計,"
        wstr2 = wstr2 & " 貸方勘定科目,貸方勘定科目名,貸方補助科目,貸方補助科目名, Sum(貸方金額) AS 貸方金額合計"
        wstr2 = wstr2 & " From DCDA040_仕訳データ2"
        wstr2 = wstr2 & " GROUP BY 番号,年月日,日番号,銀行番号,仕訳区分,仕訳補助,仕訳名,社債フラグ,借方勘定科目,借方勘定科目名,借方補助科目,借方補助科目名,貸方勘定科目,貸方勘定科目名,貸方補助科目,貸方補助科目名"
        wstr2 = wstr2 & " HAVING 仕訳区分='3'"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
        Do Until wRs2.EOF
            
                wRs.AddNew
                    
                    wRs("番号") = wRs2("番号")
                    wRs("年月日") = wRs2("年月日")
                    wRs("日番号") = wRs2("日番号")
                    wRs("銀行番号") = wRs2("銀行番号")
                    wRs("借入番号") = ""
                    wRs("仕訳区分") = wRs2("仕訳区分")
                    wRs("仕訳補助") = wRs2("仕訳補助")
                    wRs("仕訳名") = wRs2("仕訳名")
                    wRs("社債フラグ") = wRs2("社債フラグ")
    
                    wRs("借方勘定科目") = wRs2("借方勘定科目")
                    wRs("借方勘定科目名") = wRs2("借方勘定科目名")
                    wRs("借方金額") = wRs2("借方金額合計")
                    wRs("貸方勘定科目") = wRs2("貸方勘定科目")
                    wRs("貸方勘定科目名") = wRs2("貸方勘定科目名")
                    wRs("貸方金額") = wRs2("貸方金額合計")
    
                    wRs("借方補助科目") = wRs2("借方補助科目")
                    wRs("借方補助科目名") = wRs2("借方補助科目名")
                    wRs("銀行番号") = wRs2("銀行番号")
                    wRs("貸方補助科目") = wRs2("貸方補助科目")
                    wRs("貸方補助科目名") = wRs2("貸方補助科目名")
    
                wRs.Update
            
            wRs2.MoveNext
        Loop
        wRs2.Close
        Set wRs2 = Nothing
        
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA010_月次仕訳作成_日本ガス_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA010_月次仕訳作成_日本ガス() でエラー" + vbCrLf + vbCrLf + _
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
' MDA010_仕訳計上科目作成_日本ガス
'------------------------------------------------
Public Sub MDA010_仕訳計上科目作成_日本ガス()
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String

    Dim w勘定科目 As MDA010_勘定科目
    Dim wDate1 As Date, wDate2 As Date
    Dim w番号 As String
'
    On Error GoTo MDA010_仕訳計上科目作成_日本ガス_ERR
'
    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    wDate1 = C年月日.平成To西暦("年月日", GRpt.テキスト_01)
    wDate1 = MBA010_締日年月日(wDate1)
    wDate2 = DateAdd("d", 1, wDate1)
    
    w番号 = Right("00" + CStr(GInt1), 2)
'
    p仕訳データ.社債フラグ = 0
    p仕訳データ.借入番号 = ""
    p仕訳データ.銀行番号 = ""
    p仕訳データ.長短区分 = ""
    p仕訳データ.利息区分 = ""
    p仕訳データ.仕訳区分 = ""
    p仕訳データ.借方勘定科目 = ""
    p仕訳データ.貸方勘定科目 = ""

    w番号 = Right("00" & CStr(GInt1), 2)
'
    'ワークデータ 計上仕訳作成と振り戻し仕訳作成
    wstr = ""
    wstr = wstr + "Select * From DCDA040_仕訳データ"
    'wstr = wstr + "Select * From DCDA040_仕訳データ2"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        wstr2 = ""
        wstr2 = wstr2 & "SELECT K.銀行番号, K.利息区分, S.社債フラグ,"
        wstr2 = wstr2 & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Sum(Z.前払利息_" & w番号 & "),Sum(Z.未払利息_" & w番号 & ")) AS 利息残高合計"
        wstr2 = wstr2 & " FROM (DCDA010_借入残高推移表結果 AS Z"
        wstr2 = wstr2 & " INNER JOIN DBDA010_借入金 AS K"
        wstr2 = wstr2 & " ON Z.借入番号 = K.借入番号)"
        wstr2 = wstr2 & " LEFT JOIN DAAA116_借入金種別 AS S"
        wstr2 = wstr2 & " ON K.借入金種別区分 = S.借入金種別区分"
        wstr2 = wstr2 & " GROUP BY K.銀行番号, K.利息区分, S.社債フラグ"
        wstr2 = wstr2 & " HAVING IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Sum(Z.前払利息_" & w番号 & "),Sum(Z.未払利息_" & w番号 & "))<>0"
        wstr2 = wstr2 & " ORDER BY K.銀行番号"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
            Do Until wRs2.EOF
                p仕訳データ.社債フラグ = P8.FCDbl(wRs2("社債フラグ"))
                p仕訳データ.借入番号 = ""
                p仕訳データ.長短区分 = ""
                p仕訳データ.利息区分 = wRs2("利息区分")
                p仕訳データ.銀行番号 = wRs2("銀行番号")

                '利息額　支払利息/利息前払未払費用
                    '----------< MDA010_勘定科目Read >----------
                    p仕訳データ.仕訳区分 = "4"
                    w勘定科目 = MDA010_勘定科目Read()

                    If w勘定科目.借方勘定科目 <> "" Then
                        If p仕訳データ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
                            wRs.AddNew
                                wRs("番号") = 4
                                wRs("年月日") = wDate1
                                wRs("日番号") = 0
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ

                                wRs("借方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("借方金額") = wRs2("利息残高合計")
                                wRs("貸方勘定科目") = w勘定科目.借方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("貸方金額") = wRs2("利息残高合計")

                                wRs("銀行番号") = w勘定科目.借方銀行番号
                                wRs("借方補助科目") = w勘定科目.貸方補助科目
                                wRs("借方補助科目名") = w勘定科目.貸方補助科目名
                                wRs("貸方補助科目") = w勘定科目.借方補助科目
                                wRs("貸方補助科目名") = w勘定科目.借方補助科目名

                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If

                            wRs.Update

                            '振り戻し作成
                            wRs.AddNew

                                wRs("番号") = 4
                                wRs("年月日") = wDate2
                                wRs("日番号") = 0
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名 & "/振り戻し"
                                wRs("社債フラグ") = p仕訳データ.社債フラグ

                                wRs("銀行番号") = w勘定科目.借方銀行番号
                                wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                                wRs("借方補助科目") = w勘定科目.借方補助科目
                                wRs("借方補助科目名") = w勘定科目.借方補助科目名

                                wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("貸方金額") = wRs2("利息残高合計")
                                wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("借方金額") = wRs2("利息残高合計")

                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If

                            wRs.Update

                        ElseIf p仕訳データ.利息区分 = XMXA020_区分("利息区分", "利息後払") Then
                            wRs.AddNew
    
                                wRs("番号") = 4
                                wRs("年月日") = wDate1
                                wRs("日番号") = 0
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ
    
                                wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("借方金額") = wRs2("利息残高合計")
                                wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("貸方金額") = wRs2("利息残高合計")
    
                                wRs("銀行番号") = w勘定科目.借方銀行番号
                                wRs("借方補助科目") = w勘定科目.借方補助科目
                                wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
    
                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If
    
                            wRs.Update
    
                            '振り戻し作成
                            wRs.AddNew
    
                                wRs("番号") = 4
                                wRs("年月日") = wDate2
                                wRs("日番号") = 0
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名 & "/振り戻し"
                                wRs("社債フラグ") = p仕訳データ.社債フラグ
    
                                wRs("銀行番号") = w勘定科目.借方銀行番号
                                wRs("貸方補助科目") = w勘定科目.借方補助科目
                                wRs("貸方補助科目名") = w勘定科目.借方補助科目名
                                wRs("借方補助科目") = w勘定科目.貸方補助科目
                                wRs("借方補助科目名") = w勘定科目.貸方補助科目名
    
                                wRs("貸方勘定科目") = w勘定科目.借方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("貸方金額") = wRs2("利息残高合計")
                                wRs("借方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("借方金額") = wRs2("利息残高合計")
    
                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If
    
                            wRs.Update
                        End If
                    End If

                wRs2.MoveNext
            Loop
        wRs2.Close
        Set wRs2 = Nothing
    
    
    
    wRs.Close
    Set wRs = Nothing
'

'
'    '銀行 金額集計
'    wstr = ""
'    wstr = wstr + "Select * From DCDA040_仕訳データ"
'    Call AdoRecordsetOpen(GDb, wRs, wstr)
'
'        wstr2 = ""
'        wstr2 = wstr2 & "SELECT 仕訳区分,年月日,社債フラグ,番号,仕訳補助,日番号,銀行番号,仕訳名,"
'        wstr2 = wstr2 & " 借方勘定科目,借方勘定科目名,借方補助科目,借方補助科目名, Sum(DCDA040_仕訳データ2.借方金額) AS 借方金額の合計,"
'        wstr2 = wstr2 & " 貸方勘定科目,貸方勘定科目名,貸方補助科目,貸方補助科目名, Sum(DCDA040_仕訳データ2.貸方金額) AS 貸方金額の合計"
'        wstr2 = wstr2 & " From DCDA040_仕訳データ2"
'        wstr2 = wstr2 & " GROUP BY 仕訳区分,年月日,社債フラグ,番号,仕訳補助,日番号,銀行番号,仕訳名,借方勘定科目,借方勘定科目名,借方補助科目,借方補助科目名,貸方勘定科目,貸方勘定科目名,貸方補助科目,貸方補助科目名"
'        wstr2 = wstr2 & " ORDER BY 仕訳区分,年月日,社債フラグ,番号,仕訳補助,日番号"
'        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
'            Do Until wRs2.EOF
'
'                wRs.AddNew
'
'                    wRs("番号") = wRs2("番号")
'                    wRs("年月日") = wRs2("年月日")
'                    wRs("日番号") = wRs2("日番号")
'                    wRs("借入番号") = ""
'                    wRs("仕訳区分") = wRs2("仕訳区分")
'                    wRs("仕訳補助") = wRs2("仕訳補助")
'                    wRs("仕訳名") = wRs2("仕訳名")
'                    wRs("社債フラグ") = wRs2("社債フラグ")
'
'                    wRs("借方勘定科目") = wRs2("借方勘定科目")
'                    wRs("借方勘定科目名") = wRs2("借方勘定科目名")
'                    wRs("借方金額") = wRs2("借方金額の合計")
'                    wRs("貸方勘定科目") = wRs2("貸方勘定科目")
'                    wRs("貸方勘定科目名") = wRs2("貸方勘定科目名")
'                    wRs("貸方金額") = wRs2("貸方金額の合計")
'
'                    wRs("銀行番号") = wRs2("銀行番号")
'                    wRs("貸方補助科目") = wRs2("貸方補助科目")
'                    wRs("貸方補助科目名") = wRs2("貸方補助科目名")
'                    wRs("借方補助科目") = wRs2("借方補助科目")
'                    wRs("借方補助科目名") = wRs2("借方補助科目名")
'
'                wRs.Update
'
'                wRs2.MoveNext
'            Loop
'        wRs2.Close
'        Set wRs2 = Nothing
'
'    wRs.Close
'    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA010_仕訳計上科目作成_日本ガス_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA010_仕訳計上科目作成_日本ガス() でエラー" + vbCrLf + vbCrLf + _
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
' MDA010_仕訳長短振替科目作成_日本ガス
'------------------------------------------------
Public Sub MDA010_仕訳長短振替科目作成_日本ガス()
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String

    Dim w勘定科目 As MDA010_勘定科目
    Dim wDate1 As Date, wDate2 As Date
    Dim w番号 As String, w番号2 As String
'
    On Error GoTo MDA010_仕訳長短振替科目作成_日本ガス_ERR
'
    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    wDate1 = C年月日.平成To西暦("年月日", GRpt.テキスト_01)
    wDate1 = MBA010_締日年月日(wDate1)
    wDate2 = DateAdd("d", 1, wDate1)

    w番号 = Right("00" + CStr(GInt1 + 1), 2)
    w番号2 = Right("00" + CStr(GInt1 + 2), 2)
'
    p仕訳データ.社債フラグ = 0
    p仕訳データ.借入番号 = ""
    p仕訳データ.銀行番号 = ""
    p仕訳データ.長短区分 = ""
    p仕訳データ.利息区分 = ""
    p仕訳データ.仕訳区分 = ""
    p仕訳データ.借方勘定科目 = ""
    p仕訳データ.貸方勘定科目 = ""
'
'
    '銀行 金額集計/振替仕訳
    wstr = ""
    wstr = wstr + "Select * From DCDA040_仕訳データ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        wstr2 = ""
        wstr2 = wstr2 & "SELECT KS.社債フラグ,K.銀行番号,S.借入番号,"
        wstr2 = wstr2 & " K.長短区分,K.利息区分,"
        wstr2 = wstr2 & " (S.元金_" & w番号 & "+S.元金_" & w番号2 & ") AS 元金額" '1年間（半期+半期）
        wstr2 = wstr2 & " FROM (DCDA010_借入残高推移表結果 AS S"
        wstr2 = wstr2 & " INNER JOIN DBDA010_借入金 AS K"
        wstr2 = wstr2 & " ON S.借入番号 = K.借入番号)"
        wstr2 = wstr2 & " INNER JOIN DAAA116_借入金種別 KS"
        wstr2 = wstr2 & " ON K.借入金種別区分 = KS.借入金種別区分"
        wstr2 = wstr2 & " Where K.長短区分=" & P8.FCDbl(XMXA020_区分("長短区分", "長期借入金"))
        wstr2 = wstr2 & " And (S.元金_" & w番号 & " <> 0"
        wstr2 = wstr2 & " Or S.元金_" & w番号2 & " <> 0)"
        wstr2 = wstr2 & " ORDER BY KS.社債フラグ, K.銀行番号,S.借入番号"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
        Do Until wRs2.EOF
            
                p仕訳データ.社債フラグ = P8.FCDbl(wRs2("社債フラグ"))
                p仕訳データ.借入番号 = wRs2("借入番号")
                p仕訳データ.銀行番号 = wRs2("銀行番号")
                p仕訳データ.長短区分 = wRs2("長短区分")
                p仕訳データ.利息区分 = wRs2("利息区分")
                
                '----------< MDA010_勘定科目Read >----------
                p仕訳データ.仕訳区分 = "7"
                w勘定科目 = MDA010_勘定科目Read()
                
                If w勘定科目.借方勘定科目 <> "" Then
                    wRs.AddNew
                        
                        wRs("番号") = 1
                        wRs("年月日") = wDate1
                        wRs("日番号") = 0
                        wRs("借入番号") = p仕訳データ.借入番号
                        wRs("仕訳区分") = w勘定科目.仕訳区分
                        wRs("仕訳補助") = w勘定科目.仕訳補助
                        wRs("仕訳名") = w勘定科目.仕訳名
                        wRs("社債フラグ") = p仕訳データ.社債フラグ

                        wRs("借方勘定科目") = w勘定科目.借方勘定科目
                        wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                        wRs("借方金額") = wRs2("元金額")
                        wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                        wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                        wRs("貸方金額") = wRs2("元金額")

                        wRs("借方補助科目") = w勘定科目.借方補助科目
                        wRs("借方補助科目名") = w勘定科目.借方補助科目名
                        wRs("銀行番号") = w勘定科目.貸方銀行番号
                        wRs("貸方補助科目") = w勘定科目.貸方補助科目
                        wRs("貸方補助科目名") = w勘定科目.貸方補助科目名

                        If wRs("銀行番号") = "" Then
                            wRs("銀行番号") = p仕訳データ.銀行番号
                        End If
                    
                    wRs.Update
                
                    '振り戻し作成
                    wRs.AddNew

                        wRs("番号") = 1
                        wRs("年月日") = wDate2
                        wRs("日番号") = 0
                        wRs("借入番号") = p仕訳データ.借入番号
                        wRs("仕訳区分") = w勘定科目.仕訳区分
                        wRs("仕訳補助") = w勘定科目.仕訳補助
                        wRs("仕訳名") = w勘定科目.仕訳名 & "/振り戻し"
                        wRs("社債フラグ") = p仕訳データ.社債フラグ

                        wRs("貸方勘定科目") = w勘定科目.借方勘定科目
                        wRs("貸方勘定科目名") = w勘定科目.借方勘定科目名
                        wRs("貸方金額") = wRs2("元金額")
                        wRs("借方勘定科目") = w勘定科目.貸方勘定科目
                        wRs("借方勘定科目名") = w勘定科目.貸方勘定科目名
                        wRs("借方金額") = wRs2("元金額")

                        wRs("貸方補助科目") = w勘定科目.借方補助科目
                        wRs("貸方補助科目名") = w勘定科目.借方補助科目名
                        wRs("銀行番号") = w勘定科目.貸方銀行番号
                        wRs("借方補助科目") = w勘定科目.貸方補助科目
                        wRs("借方補助科目名") = w勘定科目.貸方補助科目名
                        
                        If wRs("銀行番号") = "" Then
                            wRs("銀行番号") = p仕訳データ.銀行番号
                        End If
        
                    wRs.Update
                End If
            
            wRs2.MoveNext
        Loop
        wRs2.Close
        Set wRs2 = Nothing
        
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA010_仕訳長短振替科目作成_日本ガス_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA010_仕訳長短振替科目作成_日本ガス() でエラー" + vbCrLf + vbCrLf + _
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
' MDA010_仕訳現金科目作成_神姫バス
'------------------------------------------------
Public Sub MDA010_仕訳現金科目作成_神姫バス(p借入計画マスタ As MAA910_借入金)
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String

    Dim w勘定科目 As MDA010_勘定科目
    Dim j As Integer
    Dim wDate1 As Date, wDate2 As Date
    Dim p借入金種別 As MAA070_借入金種別
    
    Dim ws利息開始日 As String, ws利息終了日 As String
    Dim wDateST As Date, wDateED As Date
    Dim w銀行マスタ As MAA030_銀行
'
    On Error GoTo MDA010_仕訳現金科目作成_神姫バス_ERR
'
    wDate1 = C年月日.平成To西暦("年月日", GRpt.テキスト_01)
    wDate1 = MBA010_締日年月日(wDate1)
    wDate1 = DateAdd("m", -1, wDate1)
    wDate1 = MBA010_締日年月日(Format(wDate1, "yyyy/mm/01"))
    
    wDate2 = C年月日.平成To西暦("年月日", GRpt.テキスト_02)
    wDate2 = MBA010_締日年月日(wDate2)
'
    p仕訳データ.借入番号 = ""
    p仕訳データ.銀行番号 = ""
    p仕訳データ.社債フラグ = 0
    p仕訳データ.長短区分 = ""
    p仕訳データ.利息区分 = ""
    p仕訳データ.仕訳区分 = ""
    p仕訳データ.借方勘定科目 = ""
    p仕訳データ.貸方勘定科目 = ""
'
    '借入金種別区分
    p借入金種別 = MAA070_借入金種別Read(p借入計画マスタ.借入金種別区分)

    '銀行マスタ
    w銀行マスタ = MAA030_銀行マスタRead(p借入計画マスタ.銀行番号)

'
    wstr = ""
    wstr = wstr & "Select * From DCDA040_仕訳データ2"
    'wstr = wstr & "Select * From DCDA040_仕訳データ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        p仕訳データ.社債フラグ = p借入金種別.社債フラグ
        p仕訳データ.借入番号 = p借入計画マスタ.借入番号
        p仕訳データ.長短区分 = p借入計画マスタ.長短区分
        p仕訳データ.利息区分 = p借入計画マスタ.利息区分
        p仕訳データ.銀行番号 = p借入計画マスタ.銀行番号
        
        '借入金の実行　現金科目/借入金
        'If wDate1 <= p借入計画マスタ.実行日 And wDate2 > p借入計画マスタ.実行日 Then
        If wDate1 < p借入計画マスタ.実行日 And wDate2 >= p借入計画マスタ.実行日 Then
            '----------< MDA010_勘定科目Read >----------
            p仕訳データ.仕訳区分 = "1"
            w勘定科目 = MDA010_勘定科目Read()
                
            If w勘定科目.借方勘定科目 <> "" Then
                wRs.AddNew
                    wRs("番号") = 0
                    wRs("年月日") = p借入計画マスタ.実行日
                    wSDate = MBA010_対象年月(CDate(p借入計画マスタ.実行日))
                    wRs("対象年月") = wSDate
                    '
                    wRs("借入番号") = p仕訳データ.借入番号
                    wRs("仕訳区分") = w勘定科目.仕訳区分
                    wRs("仕訳補助") = w勘定科目.仕訳補助
                    wRs("仕訳名") = w勘定科目.仕訳名
                    wRs("社債フラグ") = p仕訳データ.社債フラグ
                    wRs("借方勘定科目") = w勘定科目.借方勘定科目
                    wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                    wRs("借方金額") = p借入計画マスタ.融資金額
                    wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                    wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                    wRs("貸方金額") = p借入計画マスタ.融資金額
                
                    wRs("銀行番号") = w勘定科目.借方銀行番号
                    wRs("借方補助科目") = w勘定科目.借方補助科目
                    wRs("借方補助科目名") = w勘定科目.借方補助科目名
                    wRs("貸方補助科目") = w勘定科目.貸方補助科目
                    wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                
                    If wRs("銀行番号") = "" Then
                        wRs("銀行番号") = p仕訳データ.銀行番号
                    End If
                    
                    '神姫バス
                    If p仕訳データ.長短区分 = XMXA020_区分("長短区分", "長期借入金") Then
                        wRs("摘要") = "長期 借入 " & p仕訳データ.借入番号
                    Else
                        wRs("摘要") = "短期 借入 " & p仕訳データ.借入番号
                    End If
                    wRs("伝票番号") = GstrDenNo
                
                wRs.Update
            End If
        End If

        '借入金の返済　借入金/現金科目
        For j = 1 To UBound(G借入金テーブル)
            
            p仕訳データ.仕訳区分 = ""
            p仕訳データ.借方勘定科目 = ""
            p仕訳データ.貸方勘定科目 = ""
                    
            'If wDate1 <= G借入金テーブル(j).実際年月日 And wDate2 > G借入金テーブル(j).実際年月日 Then
            If wDate1 < G借入金テーブル(j).実際年月日 And wDate2 >= G借入金テーブル(j).実際年月日 Then
                If G借入金テーブル(j).元金額 <> 0 Or G借入金テーブル(j).利息額 <> 0 _
                Or (G借入金テーブル(j).融資残高 <> 0 And p借入計画マスタ.実行日 = G借入金テーブル(j).実際年月日) _
                Or G借入金テーブル(j).保証料 <> 0 Or G借入金テーブル(j).手数料 <> 0 _
                Or Format(p借入計画マスタ.解約実行日, "yyyymmdd") = Format(G借入金テーブル(j).実際年月日, "yyyymmdd") Then '10/06/16 V195
                   
                    '元金額
                    If Format(p借入計画マスタ.解約実行日, "yyyymmdd") = Format(G借入金テーブル(j).実際年月日, "yyyymmdd") _
                    And G借入金テーブル(j).融資残高 <> 0 Then
                    '解約算出
                        '----------< MDA010_勘定科目Read >----------
                        p仕訳データ.仕訳区分 = "2"
                        w勘定科目 = MDA010_勘定科目Read()
                        
                        If w勘定科目.借方勘定科目 <> "" Then
                            wRs.AddNew
                                
                                wRs("番号") = 1
                                wRs("年月日") = G借入金テーブル(j).実際年月日
                                wSDate = MBA010_対象年月(CDate(G借入金テーブル(j).実際年月日))
                                wRs("対象年月") = wSDate
                                '
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ
                                wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("借方金額") = G借入金テーブル(j).融資残高
                                wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("貸方金額") = G借入金テーブル(j).融資残高
                            
                                wRs("借方補助科目") = w勘定科目.借方補助科目
                                wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                wRs("銀行番号") = w勘定科目.貸方銀行番号
                                wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                            
                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If
                            
                                '神姫バス
                                If p仕訳データ.長短区分 = XMXA020_区分("長短区分", "長期借入金") Then
                                    wRs("摘要") = "長期 返済 " & p仕訳データ.借入番号
                                Else
                                    wRs("摘要") = "短期 返済 " & p仕訳データ.借入番号
                                End If
                                wRs("伝票番号") = GstrDenNo
                                        
                            wRs.Update
                        End If
                    ElseIf G借入金テーブル(j).元金額 <> 0 Then
                        '----------< MDA010_勘定科目Read >----------
                        p仕訳データ.仕訳区分 = "2"
                        w勘定科目 = MDA010_勘定科目Read()
                        
                        If w勘定科目.借方勘定科目 <> "" Then
                            wRs.AddNew
                                wRs("番号") = 1
                                wRs("年月日") = G借入金テーブル(j).実際年月日
                                wSDate = MBA010_対象年月(CDate(G借入金テーブル(j).実際年月日))
                                wRs("対象年月") = wSDate
                                '
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ
                                wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("借方金額") = G借入金テーブル(j).元金額
                                wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("貸方金額") = G借入金テーブル(j).元金額
                            
                                wRs("借方補助科目") = w勘定科目.借方補助科目
                                wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                wRs("銀行番号") = w勘定科目.貸方銀行番号
                                wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                            
                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If
                            
                                '神姫バス
                                If p仕訳データ.長短区分 = XMXA020_区分("長短区分", "長期借入金") Then
                                    wRs("摘要") = "長期 返済 " & p仕訳データ.借入番号
                                Else
                                    wRs("摘要") = "短期 返済 " & p仕訳データ.借入番号
                                End If
                                wRs("伝票番号") = GstrDenNo
                
                            wRs.Update
                        End If
                    End If
                    
                    If G借入金テーブル(j).利息額 <> 0 Then
                    '利息額　利息前払未払費用/普通預金
                        '----------< MDA010_勘定科目Read >----------
                        p仕訳データ.仕訳区分 = "3"
                        w勘定科目 = MDA010_勘定科目Read()
                        
                        If w勘定科目.借方勘定科目 <> "" Then
                            wRs.AddNew
                                wRs("番号") = 3
                                wRs("年月日") = G借入金テーブル(j).実際年月日
                                wSDate = MBA010_対象年月(CDate(G借入金テーブル(j).実際年月日))
                                wRs("対象年月") = wSDate
                                '
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ
                                wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("借方金額") = G借入金テーブル(j).利息額
                                wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("貸方金額") = G借入金テーブル(j).利息額
                            
                                wRs("借方補助科目") = w勘定科目.借方補助科目
                                wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                wRs("銀行番号") = w勘定科目.貸方銀行番号
                                wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                            
                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If
                                    
                                '神姫バス
                                '摘要 利息期間
                                If p借入計画マスタ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
                                    '利息開始日
                                    If p借入計画マスタ.実行日 = G借入金テーブル(j).利息計算年月日 Then
                                    '実行日
                                        If p借入計画マスタ.利息控除区分 = P8.FCDbl(XMXA020_区分("利息控除", "実行日控除")) _
                                        Or p借入計画マスタ.利息控除区分 = P8.FCDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) Then
                                        '実行日控除
                                            wDateST = DateAdd("d", 1, G借入金テーブル(j).利息計算年月日)
                                        Else '控除無し
                                            wDateST = G借入金テーブル(j).利息計算年月日
                                        End If
                                    Else
                                        wDateST = DateAdd("d", 1, G借入金テーブル(j).利息計算年月日)
                                    End If
                                
                                    '利息終了日
                                    If p借入計画マスタ.解約実行日 = G借入金テーブル(j).利息計算年月日 _
                                    Or G借入金テーブル(j).日割日数 < 0 Then
                                    '解約日 or 内入
                                        wDateED = DateAdd("d", Abs(G借入金テーブル(j).日割日数), wDateST)
                                    Else
                                        wDateED = DateAdd("d", G借入金テーブル(j).日割日数 - 1, wDateST)
                                    End If
                                
                                Else '利息後払
                                
                                    '利息終了日
                                    If p借入計画マスタ.最終返済実行日 = G借入金テーブル(j).利息計算年月日 Then
                                    '最終返済
                                        If p借入計画マスタ.利息控除区分 = P8.FCDbl(XMXA020_区分("利息控除", "最終返済日控除")) _
                                        Or p借入計画マスタ.利息控除区分 = P8.FCDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) Then
                                        '最終返済日控除
                                            wDateED = DateAdd("d", -1, G借入金テーブル(j).利息計算年月日)
                                        Else '控除無し
                                            wDateED = G借入金テーブル(j).利息計算年月日
                                        End If
                                    Else
                                        wDateED = G借入金テーブル(j).利息計算年月日
                                    End If
                                    
                                    '利息開始日
                                    wDateST = DateAdd("d", -G借入金テーブル(j).日割日数 + 1, G借入金テーブル(j).利息計算年月日)
                                End If
                                
                                ws利息開始日 = Format(wDateST, "yy/mm/dd")
                                ws利息終了日 = Format(wDateED, "yy/mm/dd")
                                
                                If p仕訳データ.長短区分 = XMXA020_区分("長短区分", "長期借入金") Then
                                    wRs("摘要") = "長期 " & ws利息開始日 & "-" & ws利息終了日 & " " & w銀行マスタ.銀行名
                                Else
                                    wRs("摘要") = "短期 " & ws利息開始日 & "-" & ws利息終了日 & " " & w銀行マスタ.銀行名
                                End If
                                wRs("伝票番号") = GstrDenNo
                            
                            wRs.Update
                        End If
                    End If
                    
                End If
            End If
                         
            'If wDate2 <= G借入金テーブル(j).実際年月日 Then
            If wDate2 < G借入金テーブル(j).実際年月日 Then
                Exit For
            End If
        Next
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA010_仕訳現金科目作成_神姫バス_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA010_仕訳現金科目作成_神姫バス() でエラー" + vbCrLf + vbCrLf + _
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
' MDA010_仕訳現金科目作成_明細TR_神姫バス
'------------------------------------------------
Public Sub MDA010_仕訳現金科目作成_明細TR_神姫バス(pTbl As String)
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String

    Dim w勘定科目 As MDA010_勘定科目
    Dim wDate1 As Date, wDate2 As Date
    Dim w借入番号 As String

    Dim ws利息開始日 As String, ws利息終了日 As String
    Dim wDateST As Date, wDateED As Date

    Dim w銀行マスタ As MAA030_銀行
'
    On Error GoTo MDA010_仕訳現金科目作成_明細TR_神姫バス_ERR
'
    wDate1 = C年月日.平成To西暦("年月日", GRpt.テキスト_01)
    wDate1 = MBA010_締日年月日(wDate1)
    wDate1 = DateAdd("m", -1, wDate1)
    wDate1 = MBA010_締日年月日(Format(wDate1, "yyyy/mm/01"))
    
    wDate2 = C年月日.平成To西暦("年月日", GRpt.テキスト_02)
    wDate2 = MBA010_締日年月日(wDate2)
'
    p仕訳データ.社債フラグ = 0
    p仕訳データ.借入番号 = ""
    p仕訳データ.銀行番号 = ""
    p仕訳データ.長短区分 = ""
    p仕訳データ.利息区分 = ""
    p仕訳データ.仕訳区分 = ""
    p仕訳データ.借方勘定科目 = ""
    p仕訳データ.貸方勘定科目 = ""
'
    wstr = ""
    wstr = wstr + "Select * From DCDA040_仕訳データ2"
    'wstr = wstr + "Select * From DCDA040_仕訳データ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        '借入金の実行　現金科目/借入金
        
        w借入番号 = ""
        
        wstr2 = "SELECT "
        wstr2 = wstr2 & "K.実行日,"
        wstr2 = wstr2 & "K.借入番号,"
        wstr2 = wstr2 & "K.銀行番号,"
        wstr2 = wstr2 & "K.長短区分,"
        wstr2 = wstr2 & "K.利息区分,"
        wstr2 = wstr2 & "K.融資金額,"
        wstr2 = wstr2 & "K.解約実行日,"
        wstr2 = wstr2 & "TR.借入番号,"
        wstr2 = wstr2 & "TR.返済回数,"
        wstr2 = wstr2 & "TR.実際年月日,"
        wstr2 = wstr2 & "TR.返済金額,"
        wstr2 = wstr2 & "TR.元金額,"
        wstr2 = wstr2 & "TR.利息額,"
        wstr2 = wstr2 & "TR.融資残高,"
        wstr2 = wstr2 & "TR.保証料,"
        wstr2 = wstr2 & "TR.手数料,"
        wstr2 = wstr2 & "S.社債フラグ"
        wstr2 = wstr2 & " From (DBDA010_借入金 As K"
        wstr2 = wstr2 & " INNER JOIN DBDA010_借入金明細TR As TR"
        wstr2 = wstr2 & " ON K.借入番号 = TR.借入番号)"
        wstr2 = wstr2 & " INNER JOIN DAAA116_借入金種別 As S"
        wstr2 = wstr2 & " ON K.借入金種別区分 = S.借入金種別区分"
        wstr2 = wstr2 & " WHERE K.手入力区分=1 AND K.取消フラグ=0"
        wstr2 = wstr2 & " AND TR.取消フラグ=0 AND TR.取消フラグ２=0"
        wstr2 = wstr2 & " Order by TR.借入番号,TR.返済回数"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
            Do Until wRs2.EOF
                If w借入番号 <> wRs2("K.借入番号") Then
                    p仕訳データ.社債フラグ = P8.FCDbl(wRs2("社債フラグ"))
                    p仕訳データ.借入番号 = wRs2("K.借入番号")
                    p仕訳データ.長短区分 = wRs2("長短区分")
                    p仕訳データ.利息区分 = wRs2("利息区分")
                    p仕訳データ.銀行番号 = wRs2("銀行番号")
                        
                    'If wDate1 <= wRs2("実行日") And wDate2 > wRs2("実行日") Then
                    If wDate1 < wRs2("実行日") And wDate2 >= wRs2("実行日") Then
                        '----------< MDA010_勘定科目Read >----------
                        p仕訳データ.仕訳区分 = "1"
                        w勘定科目 = MDA010_勘定科目Read()
                        
                        If w勘定科目.借方勘定科目 <> "" Then
                            wRs.AddNew
                                wRs("番号") = 0
                                wRs("年月日") = wRs2("実行日")
                                wSDate = MBA010_対象年月(wRs2("実行日"))
                                wRs("対象年月") = wSDate
                                '
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ
                                wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("借方金額") = wRs2("融資金額")
                                wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("貸方金額") = wRs2("融資金額")
                            
                                wRs("銀行番号") = w勘定科目.借方銀行番号
                                wRs("借方補助科目") = w勘定科目.借方補助科目
                                wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                            
                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If
                                    
                                '神姫バス
                                If p仕訳データ.長短区分 = XMXA020_区分("長短区分", "長期借入金") Then
                                    wRs("摘要") = "長期 借入 " & p仕訳データ.借入番号
                                Else
                                    wRs("摘要") = "短期 借入 " & p仕訳データ.借入番号
                                End If
                                wRs("伝票番号") = GstrDenNo
                                        
                            wRs.Update
                        End If
                    End If
                    
                End If
                
                w借入番号 = wRs2("K.借入番号")
                
                wRs2.MoveNext
            Loop
        wRs2.Close
        Set wRs2 = Nothing
        '
        '借入金の返済　借入金/現金科目
        wstr2 = "SELECT "
        wstr2 = wstr2 & "K.実行日,"
        wstr2 = wstr2 & "K.借入番号,"
        wstr2 = wstr2 & "K.銀行番号,"
        wstr2 = wstr2 & "K.長短区分,"
        wstr2 = wstr2 & "K.利息区分,"
        wstr2 = wstr2 & "K.融資金額,"
        wstr2 = wstr2 & "K.解約実行日,"
        wstr2 = wstr2 & "K.最終返済実行日,"
        wstr2 = wstr2 & "K.利息控除区分,"
        wstr2 = wstr2 & "K.手入力区分,"
        wstr2 = wstr2 & "K.日割計算区分,"
        wstr2 = wstr2 & "TR.借入番号,"
        wstr2 = wstr2 & "TR.返済回数,"
        wstr2 = wstr2 & "TR.実際年月日,"
        wstr2 = wstr2 & "TR.利息計算年月日,"
        wstr2 = wstr2 & "TR.日割日数,"
        wstr2 = wstr2 & "TR.返済金額,"
        wstr2 = wstr2 & "TR.元金額,"
        wstr2 = wstr2 & "TR.利息額,"
        wstr2 = wstr2 & "TR.融資残高,"
        wstr2 = wstr2 & "TR.保証料,"
        wstr2 = wstr2 & "TR.手数料,"
        wstr2 = wstr2 & "S.社債フラグ"
        wstr2 = wstr2 & " From (DBDA010_借入金 As K"
        wstr2 = wstr2 & " INNER JOIN DBDA010_借入金明細TR As TR"
        wstr2 = wstr2 & " ON K.借入番号 = TR.借入番号)"
        wstr2 = wstr2 & " INNER JOIN DAAA116_借入金種別 As S"
        wstr2 = wstr2 & " ON K.借入金種別区分 = S.借入金種別区分"
        wstr2 = wstr2 & " WHERE K.手入力区分=1 AND K.取消フラグ=0"
        wstr2 = wstr2 & " AND TR.取消フラグ=0 AND TR.取消フラグ２=0"
        wstr2 = wstr2 & " Order by TR.借入番号,TR.返済回数"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
            Do Until wRs2.EOF
                
                p仕訳データ.社債フラグ = P8.FCDbl(wRs2("社債フラグ"))
                p仕訳データ.借入番号 = wRs2("K.借入番号")
                p仕訳データ.長短区分 = wRs2("長短区分")
                p仕訳データ.利息区分 = wRs2("利息区分")
                p仕訳データ.銀行番号 = wRs2("銀行番号")
                
                w銀行マスタ = MAA030_銀行マスタRead(p仕訳データ.銀行番号)
                
                'If wDate1 <= wRs2("実際年月日") And wDate2 > wRs2("実際年月日") Then
                If wDate1 < wRs2("実際年月日") And wDate2 >= wRs2("実際年月日") Then
                    If wRs2("元金額") <> 0 Or wRs2("利息額") <> 0 _
                    Or wRs2("融資残高") <> 0 And wRs2("実行日") = wRs2("実際年月日") _
                    Or wRs2("保証料") <> 0 Or wRs2("手数料") <> 0 _
                    Or Format(wRs2("解約実行日"), "yyyymmdd") = Format(wRs2("実際年月日"), "yyyymmdd") Then '10/06/16 V195
                        
                        '元金額
                        If Format(wRs2("解約実行日"), "yyyymmdd") = Format(wRs2("実際年月日"), "yyyymmdd") _
                        And wRs2("融資残高") <> 0 Then
                        '解約算出
                            '----------< MDA010_勘定科目Read >----------
                            p仕訳データ.仕訳区分 = "2"
                            w勘定科目 = MDA010_勘定科目Read()
                            
                            If w勘定科目.借方勘定科目 <> "" Then
                                wRs.AddNew
                                    wRs("番号") = 1
                                    wRs("年月日") = wRs2("実際年月日")
                                    wSDate = MBA010_対象年月(wRs2("実際年月日"))
                                    wRs("対象年月") = wSDate
                                    '
                                    wRs("借入番号") = p仕訳データ.借入番号
                                    wRs("仕訳区分") = w勘定科目.仕訳区分
                                    wRs("仕訳補助") = w勘定科目.仕訳補助
                                    wRs("仕訳名") = w勘定科目.仕訳名
                                    wRs("社債フラグ") = p仕訳データ.社債フラグ
                                    wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                    wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                    wRs("借方金額") = wRs2("融資残高")
                                    wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                    wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                    wRs("貸方金額") = wRs2("融資残高")
                                
                                    wRs("借方補助科目") = w勘定科目.借方補助科目
                                    wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                    wRs("銀行番号") = w勘定科目.貸方銀行番号
                                    wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                    wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                                
                                    If wRs("銀行番号") = "" Then
                                        wRs("銀行番号") = p仕訳データ.銀行番号
                                    End If
                                            
                                    '神姫バス
                                    If p仕訳データ.長短区分 = XMXA020_区分("長短区分", "長期借入金") Then
                                        wRs("摘要") = "長期 返済 " & p仕訳データ.借入番号
                                    Else
                                        wRs("摘要") = "短期 返済 " & p仕訳データ.借入番号
                                    End If
                                    wRs("伝票番号") = GstrDenNo
                                    
                                wRs.Update
                            End If
                        
                        ElseIf wRs2("元金額") <> 0 Then
                            '----------< MDA010_勘定科目Read >----------
                            p仕訳データ.仕訳区分 = "2"
                            w勘定科目 = MDA010_勘定科目Read()
                            
                            If w勘定科目.借方勘定科目 <> "" Then
                                wRs.AddNew
                                    wRs("番号") = 1
                                    wRs("年月日") = wRs2("実際年月日")
                                    wSDate = MBA010_対象年月(wRs2("実際年月日"))
                                    wRs("対象年月") = wSDate
                                    '
                                    wRs("借入番号") = p仕訳データ.借入番号
                                    wRs("仕訳区分") = w勘定科目.仕訳区分
                                    wRs("仕訳補助") = w勘定科目.仕訳補助
                                    wRs("仕訳名") = w勘定科目.仕訳名
                                    wRs("社債フラグ") = p仕訳データ.社債フラグ
                                    wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                    wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                    wRs("借方金額") = wRs2("元金額")
                                    wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                    wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                    wRs("貸方金額") = wRs2("元金額")
                                
                                    wRs("借方補助科目") = w勘定科目.借方補助科目
                                    wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                    wRs("銀行番号") = w勘定科目.貸方銀行番号
                                    wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                    wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                                
                                    If wRs("銀行番号") = "" Then
                                        wRs("銀行番号") = p仕訳データ.銀行番号
                                    End If
                                            
                                    '神姫バス
                                    If p仕訳データ.長短区分 = XMXA020_区分("長短区分", "長期借入金") Then
                                        wRs("摘要") = "長期 返済 " & p仕訳データ.借入番号
                                    Else
                                        wRs("摘要") = "短期 返済 " & p仕訳データ.借入番号
                                    End If
                                    wRs("伝票番号") = GstrDenNo
                                    
                                wRs.Update
                            End If
                        End If
                        
                        If wRs2("利息額") <> 0 Then
                        '利息額　利息前払未払費用/普通預金
                            '----------< MDA010_勘定科目Read >----------
                            p仕訳データ.仕訳区分 = "3"
                            w勘定科目 = MDA010_勘定科目Read()
                            
                            If w勘定科目.借方勘定科目 <> "" Then
                                wRs.AddNew
                                
                                    wRs("番号") = 3
                                    wRs("年月日") = wRs2("実際年月日")
                                    wSDate = MBA010_対象年月(wRs2("実際年月日"))
                                    wRs("対象年月") = wSDate
                                    '
                                    wRs("借入番号") = p仕訳データ.借入番号
                                    wRs("仕訳区分") = w勘定科目.仕訳区分
                                    wRs("仕訳補助") = w勘定科目.仕訳補助
                                    wRs("仕訳名") = w勘定科目.仕訳名
                                    wRs("社債フラグ") = p仕訳データ.社債フラグ
                                    wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                    wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                    wRs("借方金額") = wRs2("利息額")
                                    wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                    wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                    wRs("貸方金額") = wRs2("利息額")
                                
                                    wRs("借方補助科目") = w勘定科目.借方補助科目
                                    wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                    wRs("銀行番号") = w勘定科目.貸方銀行番号
                                    wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                    wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                                
                                    If wRs("銀行番号") = "" Then
                                        wRs("銀行番号") = p仕訳データ.銀行番号
                                    End If
                                            
                                    '神姫バス
                                    '摘要 利息期間
                                    If wRs2("手入力区分") = XMXA020_区分("登録方法", "入力登録") And wRs2("日割計算区分") = XMXA020_区分("日割計算区分", "自動計算") Then
                                        If wRs2("利息区分") = XMXA020_区分("利息区分", "利息先払") Then
                                            '利息開始日
                                            If wRs2("実行日") = wRs2("利息計算年月日") Then
                                            '実行日
                                                If wRs2("利息控除区分") = P8.FCDbl(XMXA020_区分("利息控除", "実行日控除")) _
                                                Or wRs2("利息控除区分") = P8.FCDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) Then
                                                '実行日控除
                                                    wDateST = DateAdd("d", 1, wRs2("利息計算年月日"))
                                                Else '控除無し
                                                    wDateST = wRs2("利息計算年月日")
                                                End If
                                            Else
                                                wDateST = DateAdd("d", 1, wRs2("利息計算年月日"))
                                            End If
                                        
                                            '利息終了日
                                            If wRs2("解約実行日") = wRs2("利息計算年月日") _
                                            Or wRs2("日割日数") < 0 Then
                                            '解約日 or 内入
                                                wDateED = DateAdd("d", Abs(wRs2("日割日数")), wDateST)
                                            Else
                                                wDateED = DateAdd("d", wRs2("日割日数") - 1, wDateST)
                                            End If
                                        
                                        Else '利息後払
                                        
                                            '利息終了日
                                            If wRs2("最終返済実行日") = wRs2("利息計算年月日") Then
                                            '最終返済
                                                If wRs2("利息控除区分") = P8.FCDbl(XMXA020_区分("利息控除", "最終返済日控除")) _
                                                Or wRs2("利息控除区分") = P8.FCDbl(XMXA020_区分("利息控除", "実行日及び最終返済日控除")) Then
                                                '最終返済日控除
                                                    wDateED = DateAdd("d", -1, wRs2("利息計算年月日"))
                                                Else '控除無し
                                                    wDateED = wRs2("利息計算年月日")
                                                End If
                                            Else
                                                wDateED = wRs2("利息計算年月日")
                                            End If
                                            
                                            '利息開始日
                                            wDateST = DateAdd("d", -wRs2("日割日数") + 1, wRs2("利息計算年月日"))
                                        End If
                                    
                                        ws利息開始日 = Format(wDateST, "yy/mm/dd")
                                        ws利息終了日 = Format(wDateED, "yy/mm/dd")
                                        
                                        If wRs2("長短区分") = XMXA020_区分("長短区分", "長期借入金") Then
                                            wRs("摘要") = "長期 " & ws利息開始日 & "-" & ws利息終了日 & " " & w銀行マスタ.銀行名
                                        Else
                                            wRs("摘要") = "短期 " & ws利息開始日 & "-" & ws利息終了日 & " " & w銀行マスタ.銀行名
                                        End If
                                        wRs("伝票番号") = GstrDenNo
            
                                    Else
                                        If wRs2("長短区分") = XMXA020_区分("長短区分", "長期借入金") Then
                                            wRs("摘要") = "長期" & " " & w銀行マスタ.銀行名
                                        Else
                                            wRs("摘要") = "短期" & " " & w銀行マスタ.銀行名
                                        End If
                                        wRs("伝票番号") = GstrDenNo
                                    End If
                                
                                wRs.Update
                            
                            End If
                        End If
                        
                    End If
                  
                End If
                
'                If wDate2 <= wRs2("実際年月日") Then
'                    Exit Do
'                End If
                
                wRs2.MoveNext
            Loop
        wRs2.Close
        Set wRs2 = Nothing
        '
        '借入金の返済　借入金/現金科目
        wstr2 = "SELECT "
        wstr2 = wstr2 & "K.実行日,"
        wstr2 = wstr2 & "K.借入番号,"
        wstr2 = wstr2 & "K.銀行番号,"
        wstr2 = wstr2 & "K.長短区分,"
        wstr2 = wstr2 & "K.利息区分,"
        wstr2 = wstr2 & "TR2.借入番号,"
        wstr2 = wstr2 & "TR2.実際年月日,"
        wstr2 = wstr2 & "TR2.保証料,"
        wstr2 = wstr2 & "TR2.初期手数料+TR2.元金手数料+TR2.利息手数料 As 手数料,"
        wstr2 = wstr2 & "S.社債フラグ"
        wstr2 = wstr2 & " From (DBDA010_借入金 As K"
        wstr2 = wstr2 & " INNER JOIN DBDA010_借入金明細TR2 As TR2"
        wstr2 = wstr2 & " ON K.借入番号 = TR2.借入番号)"
        wstr2 = wstr2 & " INNER JOIN DAAA116_借入金種別 As S"
        wstr2 = wstr2 & " ON K.借入金種別区分 = S.借入金種別区分"
        wstr2 = wstr2 & " WHERE S.社債フラグ=1 AND K.取消フラグ=0"
        wstr2 = wstr2 & " And K.手入力区分=1 AND K.取消フラグ=0"
        wstr2 = wstr2 & " AND TR2.取消フラグ=0 AND TR2.取消フラグ２=0"
        wstr2 = wstr2 & " Order by TR2.借入番号,TR2.実際年月日"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
            Do Until wRs2.EOF
                p仕訳データ.社債フラグ = P8.FCDbl(wRs2("社債フラグ"))
                p仕訳データ.借入番号 = wRs2("K.借入番号")
                p仕訳データ.長短区分 = wRs2("長短区分")
                p仕訳データ.利息区分 = wRs2("利息区分")
                p仕訳データ.銀行番号 = wRs2("銀行番号")
    
                'If wDate1 <= wRs2("実際年月日") And wDate2 > wRs2("実際年月日") Then
                If wDate1 < wRs2("実際年月日") And wDate2 >= wRs2("実際年月日") Then
                    If wRs2("手数料") <> 0 Then
                    '手数料　手数料/普通預金
                        '----------< MDA010_勘定科目Read >----------
                        p仕訳データ.仕訳区分 = "5"
                        w勘定科目 = MDA010_勘定科目Read()
                        
                        If w勘定科目.借方勘定科目 <> "" Then
                            wRs.AddNew
                                wRs("番号") = 3
                                wRs("年月日") = wRs2("実際年月日")
                                wSDate = MBA010_対象年月(wRs2("実際年月日"))
                                wRs("対象年月") = wSDate
                                '
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ
                                wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("借方金額") = wRs2("手数料")
                                wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("貸方金額") = wRs2("手数料")
                            
                                wRs("借方補助科目") = w勘定科目.借方補助科目
                                wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                wRs("銀行番号") = w勘定科目.貸方銀行番号
                                wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                            
                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If
                                    
                                '神姫バス
                                If p仕訳データ.長短区分 = XMXA020_区分("長短区分", "長期借入金") Then
                                    wRs("摘要") = "長期 借入 " & p仕訳データ.借入番号
                                Else
                                    wRs("摘要") = "短期 借入 " & p仕訳データ.借入番号
                                End If
                                wRs("伝票番号") = GstrDenNo
                            
                            wRs.Update
                        End If
                    End If
                    
                    If wRs2("保証料") <> 0 Then
                    '保証料　保証料/普通預金
                        '----------< MDA010_勘定科目Read >----------
                        p仕訳データ.仕訳区分 = "6"
                        w勘定科目 = MDA010_勘定科目Read()
                        
                        If w勘定科目.借方勘定科目 <> "" Then
                            wRs.AddNew
                                wRs("番号") = 3
                                wRs("年月日") = wRs2("実際年月日")
                                wSDate = MBA010_対象年月(wRs2("実際年月日"))
                                wRs("対象年月") = wSDate
                                '
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ
                                wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("借方金額") = wRs2("保証料")
                                wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("貸方金額") = wRs2("保証料")
                            
                                wRs("借方補助科目") = w勘定科目.借方補助科目
                                wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                wRs("銀行番号") = w勘定科目.貸方銀行番号
                                wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                            
                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If
                                    
                                '神姫バス
                                If p仕訳データ.長短区分 = XMXA020_区分("長短区分", "長期借入金") Then
                                    wRs("摘要") = "長期 借入 " & p仕訳データ.借入番号
                                Else
                                    wRs("摘要") = "短期 借入 " & p仕訳データ.借入番号
                                End If
                                wRs("伝票番号") = GstrDenNo
                            
                            wRs.Update
                        End If
                    End If
                End If
                    
                wRs2.MoveNext
            Loop
        wRs2.Close
        Set wRs2 = Nothing
        '
    
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA010_仕訳現金科目作成_明細TR_神姫バス_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA010_仕訳現金科目作成_明細TR_神姫バス() でエラー" + vbCrLf + vbCrLf + _
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
' MDA010_仕訳計上科目作成_神姫バス
'------------------------------------------------
Public Sub MDA010_仕訳計上科目作成_神姫バス()
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String

    Dim w勘定科目 As MDA010_勘定科目
    Dim w銀行マスタ As MAA030_銀行
    Dim wDate1 As Date, wDate2 As Date
    Dim w番号 As String
'
    On Error GoTo MDA010_仕訳計上科目作成_神姫バス_ERR
'
    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    wDate1 = C年月日.平成To西暦("年月日", GRpt.テキスト_01)
    wDate1 = MBA010_締日年月日(wDate1)
    wDate2 = DateAdd("d", 1, wDate1)
    
    w番号 = Right("00" + CStr(GInt1), 2)
'
    p仕訳データ.社債フラグ = 0
    p仕訳データ.借入番号 = ""
    p仕訳データ.銀行番号 = ""
    p仕訳データ.長短区分 = ""
    p仕訳データ.利息区分 = ""
    p仕訳データ.仕訳区分 = ""
    p仕訳データ.借方勘定科目 = ""
    p仕訳データ.貸方勘定科目 = ""

    w番号 = Right("00" & CStr(GInt1), 2)
'
    'ワークデータ 計上仕訳作成と振り戻し仕訳作成
    wstr = ""
    wstr = wstr + "Select * From DCDA040_仕訳データ2"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        wstr2 = ""
        wstr2 = wstr2 & "SELECT Z.借入番号,K.銀行番号, K.利息区分,K.長短区分, S.社債フラグ,"
        wstr2 = wstr2 & " IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息_" & w番号 & ",Z.未払利息_" & w番号 & ") AS 利息残高"
        wstr2 = wstr2 & " FROM (DCDA010_借入残高推移表結果 AS Z"
        wstr2 = wstr2 & " INNER JOIN DBDA010_借入金 AS K"
        wstr2 = wstr2 & " ON Z.借入番号 = K.借入番号)"
        wstr2 = wstr2 & " LEFT JOIN DAAA116_借入金種別 AS S"
        wstr2 = wstr2 & " ON K.借入金種別区分 = S.借入金種別区分"
        wstr2 = wstr2 & " Where IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息_" & w番号 & ",Z.未払利息_" & w番号 & ")<>0"
        wstr2 = wstr2 & " ORDER BY Z.借入番号"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
            Do Until wRs2.EOF
                p仕訳データ.社債フラグ = P8.FCDbl(wRs2("社債フラグ"))
                p仕訳データ.借入番号 = wRs2("借入番号")
                p仕訳データ.長短区分 = wRs2("利息区分")
                p仕訳データ.利息区分 = wRs2("利息区分")
                p仕訳データ.銀行番号 = wRs2("銀行番号")

                '銀行マスタ
                w銀行マスタ = MAA030_銀行マスタRead(p仕訳データ.銀行番号)
                
                '利息額　支払利息/利息前払未払費用
                    '----------< MDA010_勘定科目Read >----------
                    p仕訳データ.仕訳区分 = "4"
                    w勘定科目 = MDA010_勘定科目Read()

                    If w勘定科目.借方勘定科目 <> "" Then
                        If p仕訳データ.利息区分 = XMXA020_区分("利息区分", "利息先払") Then
                            wRs.AddNew
                                wRs("番号") = 4
                                wRs("年月日") = wDate1
                                wSDate = MBA010_対象年月(wDate1)
                                wRs("対象年月") = wSDate
                                '
                                wRs("日番号") = 0
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ

                                wRs("借方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("借方金額") = wRs2("利息残高")
                                wRs("貸方勘定科目") = w勘定科目.借方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("貸方金額") = wRs2("利息残高")

                                wRs("銀行番号") = w勘定科目.借方銀行番号
                                wRs("借方補助科目") = w勘定科目.貸方補助科目
                                wRs("借方補助科目名") = w勘定科目.貸方補助科目名
                                wRs("貸方補助科目") = w勘定科目.借方補助科目
                                wRs("貸方補助科目名") = w勘定科目.借方補助科目名

                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If

                                wRs("摘要") = p仕訳データ.借入番号 & " " & w銀行マスタ.銀行名
                                wRs("伝票番号") = GstrDenNo
                            
                            wRs.Update
                        
                            '振り戻し
                            wRs.AddNew
                                wRs("番号") = 4
                                wRs("年月日") = wDate2
                                wSDate = MBA010_対象年月(wDate2)
                                wRs("対象年月") = wSDate
                                '
                                wRs("日番号") = 0
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ

                                wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("貸方金額") = wRs2("利息残高")
                                wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("借方金額") = wRs2("利息残高")

                                wRs("銀行番号") = w勘定科目.借方銀行番号
                                wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
                                wRs("借方補助科目") = w勘定科目.借方補助科目
                                wRs("借方補助科目名") = w勘定科目.借方補助科目名

                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If

                                wRs("摘要") = p仕訳データ.借入番号 & " " & w銀行マスタ.銀行名
                                wRs("伝票番号") = GstrDenNo2
                            
                            wRs.Update
                        
                        ElseIf p仕訳データ.利息区分 = XMXA020_区分("利息区分", "利息後払") Then
                            wRs.AddNew
    
                                wRs("番号") = 4
                                wRs("年月日") = wDate1
                                wSDate = MBA010_対象年月(wDate1)
                                wRs("対象年月") = wSDate
                                '
                                wRs("日番号") = 0
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ
    
                                wRs("借方勘定科目") = w勘定科目.借方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("借方金額") = wRs2("利息残高")
                                wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("貸方金額") = wRs2("利息残高")
    
                                wRs("銀行番号") = w勘定科目.借方銀行番号
                                wRs("借方補助科目") = w勘定科目.借方補助科目
                                wRs("借方補助科目名") = w勘定科目.借方補助科目名
                                wRs("貸方補助科目") = w勘定科目.貸方補助科目
                                wRs("貸方補助科目名") = w勘定科目.貸方補助科目名
    
                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If
    
                                wRs("摘要") = p仕訳データ.借入番号 & " " & w銀行マスタ.銀行名
                                wRs("伝票番号") = GstrDenNo
                            
                            wRs.Update
    
                            '振り戻し
                            wRs.AddNew
    
                                wRs("番号") = 4
                                wRs("年月日") = wDate2
                                wSDate = MBA010_対象年月(wDate2)
                                wRs("対象年月") = wSDate
                                '
                                wRs("日番号") = 0
                                wRs("借入番号") = p仕訳データ.借入番号
                                wRs("仕訳区分") = w勘定科目.仕訳区分
                                wRs("仕訳補助") = w勘定科目.仕訳補助
                                wRs("仕訳名") = w勘定科目.仕訳名
                                wRs("社債フラグ") = p仕訳データ.社債フラグ
    
                                wRs("貸方勘定科目") = w勘定科目.借方勘定科目
                                wRs("貸方勘定科目名") = w勘定科目.借方勘定科目名
                                wRs("貸方金額") = wRs2("利息残高")
                                wRs("借方勘定科目") = w勘定科目.貸方勘定科目
                                wRs("借方勘定科目名") = w勘定科目.貸方勘定科目名
                                wRs("借方金額") = wRs2("利息残高")
    
                                wRs("銀行番号") = w勘定科目.借方銀行番号
                                wRs("貸方補助科目") = w勘定科目.借方補助科目
                                wRs("貸方補助科目名") = w勘定科目.借方補助科目名
                                wRs("借方補助科目") = w勘定科目.貸方補助科目
                                wRs("借方補助科目名") = w勘定科目.貸方補助科目名
    
                                If wRs("銀行番号") = "" Then
                                    wRs("銀行番号") = p仕訳データ.銀行番号
                                End If
    
                                wRs("摘要") = p仕訳データ.借入番号 & " " & w銀行マスタ.銀行名
                                wRs("伝票番号") = GstrDenNo2
                            
                            wRs.Update
                        
                        End If
                    End If

                wRs2.MoveNext
            Loop
        wRs2.Close
        Set wRs2 = Nothing
   
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA010_仕訳計上科目作成_神姫バス_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA010_仕訳計上科目作成_神姫バス() でエラー" + vbCrLf + vbCrLf + _
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
' MDA010_仕訳長短振替科目作成_神姫バス
'------------------------------------------------
Public Sub MDA010_仕訳長短振替科目作成_神姫バス()
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String

    Dim w勘定科目 As MDA010_勘定科目
    Dim w銀行マスタ As MAA030_銀行
    Dim wDate1 As Date, wDate2 As Date
    Dim j As Integer, wiCnt As Integer, w間隔 As Integer
    Dim w番号 As String, ws01 As String, ws02 As String
'
    On Error GoTo MDA010_仕訳長短振替科目作成_神姫バス_ERR
'
    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    wDate1 = C年月日.平成To西暦("年月日", GRpt.テキスト_01)
    wDate1 = MBA010_締日年月日(wDate1)
    wDate2 = DateAdd("d", 1, wDate1)
'
    Select Case G基本情報.決算サイクル
    Case 1
    '月次決算
        '仮の決算月を指定付に設定し、G基本情報.決算サイクル=年次と同処理をする
        w間隔 = 12
    Case 3
        w間隔 = G基本情報.決算サイクル
    Case 6
        w間隔 = G基本情報.決算サイクル
    Case Else
        w間隔 = 12
    End Select
    
    wiCnt = 12 / w間隔
    wiCnt = GInt1 + wiCnt - 1
    
    If wiCnt > 12 Then
        wiCnt = 12
    End If
'
    p仕訳データ.社債フラグ = 0
    p仕訳データ.借入番号 = ""
    p仕訳データ.銀行番号 = ""
    p仕訳データ.長短区分 = ""
    p仕訳データ.利息区分 = ""
    p仕訳データ.仕訳区分 = ""
    p仕訳データ.借方勘定科目 = ""
    p仕訳データ.貸方勘定科目 = ""
'
    '銀行 金額集計/振替仕訳
    wstr = ""
    wstr = wstr + "Select * From DCDA040_仕訳データ2"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        wstr2 = ""
        wstr2 = wstr2 & "SELECT KS.社債フラグ,K.銀行番号,S.借入番号,"
        wstr2 = wstr2 & " K.長短区分,K.利息区分,K.実行日,"
        'wstr2 = wstr2 & " (S.元金_" & w番号 & "+S.元金_" & w番号2 & ") AS 元金額" '1年間（半期+半期)日本ガス
        
            ws01 = ""
            For j = GInt1 To wiCnt - 1
                w番号 = Right("00" + CStr(j), 2)
                ws01 = ws01 & "S.元金_" & w番号 & "+"
            Next j
            
            w番号 = Right("00" + CStr(wiCnt), 2)
            ws01 = ws01 & "S.元金_" & w番号
        
        wstr2 = wstr2 & " (" & ws01 & ") AS 元金額" '1年間
            
        wstr2 = wstr2 & " FROM (DCDA010_借入残高推移表結果 AS S"
        wstr2 = wstr2 & " INNER JOIN DBDA010_借入金 AS K"
        wstr2 = wstr2 & " ON S.借入番号 = K.借入番号)"
        wstr2 = wstr2 & " INNER JOIN DAAA116_借入金種別 KS"
        wstr2 = wstr2 & " ON K.借入金種別区分 = KS.借入金種別区分"
        wstr2 = wstr2 & " Where K.長短区分=" & P8.FCDbl(XMXA020_区分("長短区分", "長期借入金"))
        'wstr2 = wstr2 & " And (S.元金_" & w番号 & " <> 0 Or S.元金_" & w番号2 & " <> 0)"'1年間（半期+半期)日本ガス
        
        '実行日
        wstr2 = wstr2 & " And format(K.実行日,'yyyy/mm/dd')<'" & Format(wDate2, "yyyy/mm/dd") & "'"
        
            ws02 = ""
            For j = GInt1 To wiCnt - 1
                w番号 = Right("00" + CStr(j), 2)
                ws02 = ws02 & "S.元金_" & w番号 & "<> 0 Or "
            Next j
            
            w番号 = Right("00" + CStr(wiCnt), 2)
            ws02 = ws02 & "S.元金_" & w番号 & "<> 0"
        
        wstr2 = wstr2 & " And (" & ws02 & ")" '1年間
        
        wstr2 = wstr2 & " ORDER BY KS.社債フラグ, K.銀行番号,S.借入番号"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
        Do Until wRs2.EOF
            
                p仕訳データ.社債フラグ = P8.FCDbl(wRs2("社債フラグ"))
                p仕訳データ.借入番号 = wRs2("借入番号")
                p仕訳データ.銀行番号 = wRs2("銀行番号")
                p仕訳データ.長短区分 = wRs2("長短区分")
                p仕訳データ.利息区分 = wRs2("利息区分")
                
                '銀行マスタ
                w銀行マスタ = MAA030_銀行マスタRead(p仕訳データ.銀行番号)
                        
                '----------< MDA010_勘定科目Read >----------
                p仕訳データ.仕訳区分 = "7"
                w勘定科目 = MDA010_勘定科目Read()
                
                If w勘定科目.借方勘定科目 <> "" Then
                    wRs.AddNew
                        
                        wRs("番号") = 1
                        wRs("年月日") = wDate1
                        wSDate = MBA010_対象年月(wDate1)
                        wRs("対象年月") = wSDate
                        '
                        wRs("日番号") = 0
                        wRs("借入番号") = p仕訳データ.借入番号
                        wRs("仕訳区分") = w勘定科目.仕訳区分
                        wRs("仕訳補助") = w勘定科目.仕訳補助
                        wRs("仕訳名") = w勘定科目.仕訳名
                        wRs("社債フラグ") = p仕訳データ.社債フラグ

                        wRs("借方勘定科目") = w勘定科目.借方勘定科目
                        wRs("借方勘定科目名") = w勘定科目.借方勘定科目名
                        wRs("借方金額") = wRs2("元金額")
                        wRs("貸方勘定科目") = w勘定科目.貸方勘定科目
                        wRs("貸方勘定科目名") = w勘定科目.貸方勘定科目名
                        wRs("貸方金額") = wRs2("元金額")

                        wRs("借方補助科目") = w勘定科目.借方補助科目
                        wRs("借方補助科目名") = w勘定科目.借方補助科目名
                        wRs("銀行番号") = w勘定科目.貸方銀行番号
                        wRs("貸方補助科目") = w勘定科目.貸方補助科目
                        wRs("貸方補助科目名") = w勘定科目.貸方補助科目名

                        If wRs("銀行番号") = "" Then
                            wRs("銀行番号") = p仕訳データ.銀行番号
                        End If
                    
                        wRs("摘要") = p仕訳データ.借入番号 & " " & w銀行マスタ.銀行名
                        wRs("伝票番号") = GstrDenNo
                    
                    wRs.Update
                
                    '振り戻し作成
                    wRs.AddNew

                        wRs("番号") = 1
                        wRs("年月日") = wDate2
                        wSDate = MBA010_対象年月(wDate2)
                        wRs("対象年月") = wSDate
                        '
                        wRs("日番号") = 0
                        wRs("借入番号") = p仕訳データ.借入番号
                        wRs("仕訳区分") = w勘定科目.仕訳区分
                        wRs("仕訳補助") = w勘定科目.仕訳補助
                        wRs("仕訳名") = w勘定科目.仕訳名
                        wRs("社債フラグ") = p仕訳データ.社債フラグ

                        wRs("貸方勘定科目") = w勘定科目.借方勘定科目
                        wRs("貸方勘定科目名") = w勘定科目.借方勘定科目名
                        wRs("貸方金額") = wRs2("元金額")
                        wRs("借方勘定科目") = w勘定科目.貸方勘定科目
                        wRs("借方勘定科目名") = w勘定科目.貸方勘定科目名
                        wRs("借方金額") = wRs2("元金額")

                        wRs("貸方補助科目") = w勘定科目.借方補助科目
                        wRs("貸方補助科目名") = w勘定科目.借方補助科目名
                        wRs("銀行番号") = w勘定科目.貸方銀行番号
                        wRs("借方補助科目") = w勘定科目.貸方補助科目
                        wRs("借方補助科目名") = w勘定科目.貸方補助科目名

                        If wRs("銀行番号") = "" Then
                            wRs("銀行番号") = p仕訳データ.銀行番号
                        End If
                    
                        wRs("摘要") = p仕訳データ.借入番号 & " " & w銀行マスタ.銀行名
                        wRs("伝票番号") = GstrDenNo2
        
                    wRs.Update
                
                End If
            
            wRs2.MoveNext
        Loop
        wRs2.Close
        Set wRs2 = Nothing
        
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA010_仕訳長短振替科目作成_神姫バス_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA010_仕訳長短振替科目作成_神姫バス() でエラー" + vbCrLf + vbCrLf + _
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
' MDA010_長短振替表_神姫バス
'------------------------------------------------
Public Sub MDA010_長短振替表_神姫バス()
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String

    Dim w勘定科目 As MDA010_勘定科目
    Dim wDate1 As Date, wDate2 As Date
    Dim j As Integer, wiCnt As Integer, w間隔 As Integer
    Dim w番号 As String, ws01 As String, ws02 As String
'
    On Error GoTo MDA010_長短振替表_神姫バス_ERR
'
    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    wDate1 = C年月日.平成To西暦("年月日", GRpt.テキスト_01)
    wDate1 = MBA010_締日年月日(wDate1)
    wDate2 = DateAdd("d", 1, wDate1)
'
    Select Case G基本情報.決算サイクル
    Case 1
    '月次決算
        '仮の決算月を指定付に設定し、G基本情報.決算サイクル=年次と同処理をする
        w間隔 = 12
    Case 3
        w間隔 = G基本情報.決算サイクル
    Case 6
        w間隔 = G基本情報.決算サイクル
    Case Else
        w間隔 = 12
    End Select
    
    wiCnt = 12 / w間隔
    wiCnt = GInt1 + wiCnt - 1
    
    If wiCnt > 12 Then
        wiCnt = 12
    End If
'
    p仕訳データ.社債フラグ = 0
    p仕訳データ.借入番号 = ""
    p仕訳データ.銀行番号 = ""
    p仕訳データ.長短区分 = ""
    p仕訳データ.利息区分 = ""
    p仕訳データ.仕訳区分 = ""
    p仕訳データ.借方勘定科目 = ""
    p仕訳データ.貸方勘定科目 = ""
'
    '銀行 金額集計/振替仕訳
    wstr = ""
    wstr = wstr + "Select * From DCKA010_資金繰表"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        wstr2 = ""
        wstr2 = wstr2 & "SELECT KS.社債フラグ,K.銀行番号,S.借入番号,"
        wstr2 = wstr2 & " K.長短区分,K.利息区分,"
        
        w番号 = Right("00" + CStr(GInt2), 2)
        wstr2 = wstr2 & "S.残高_" & w番号 & " As 基準月融資残高,"
        
            ws01 = ""
            For j = GInt1 To wiCnt - 1
                w番号 = Right("00" + CStr(j), 2)
                ws01 = ws01 & "S.元金_" & w番号 & "+"
            Next j
            
            w番号 = Right("00" + CStr(wiCnt), 2)
            ws01 = ws01 & "S.元金_" & w番号
        
        wstr2 = wstr2 & " (" & ws01 & ") AS 元金額" '1年間
            
        wstr2 = wstr2 & " FROM (DCDA010_借入残高推移表結果 AS S"
        wstr2 = wstr2 & " INNER JOIN DBDA010_借入金 AS K"
        wstr2 = wstr2 & " ON S.借入番号 = K.借入番号)"
        wstr2 = wstr2 & " INNER JOIN DAAA116_借入金種別 KS"
        wstr2 = wstr2 & " ON K.借入金種別区分 = KS.借入金種別区分"
        'wstr2 = wstr2 & " Where K.長短区分=" & P8.FCDbl(XMXA020_区分("長短区分", "長期借入金"))
        
        w番号 = Right("00" + CStr(GInt2), 2)
        wstr2 = wstr2 & " Where S.残高_" & w番号 & "<>0 "
        
        'wstr2 = wstr2 & " Where 1=1 "
            
            ws02 = ""
            For j = GInt1 To wiCnt - 1
                w番号 = Right("00" + CStr(j), 2)
                ws02 = ws02 & "S.元金_" & w番号 & "<> 0 Or "
            Next j
            
            w番号 = Right("00" + CStr(wiCnt), 2)
            ws02 = ws02 & "S.元金_" & w番号 & "<> 0"
        
        wstr2 = wstr2 & " And (" & ws02 & ")" '1年間
        
        wstr2 = wstr2 & " ORDER BY KS.社債フラグ, K.銀行番号,S.借入番号"
        Call AdoRecordsetOpen(GDb, wRs2, wstr2)
        Do Until wRs2.EOF
            
                wRs.AddNew
            
                    wRs("借入番号") = wRs2("借入番号")
                    'wRs("実際年月日") = wRs2("借入番号")
                    wRs("元金額") = wRs2("元金額")
                    wRs("融資残高") = wRs2("基準月融資残高")
                
                wRs.Update
            
            wRs2.MoveNext
        Loop
        wRs2.Close
        Set wRs2 = Nothing
        
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MDA010_長短振替表_神姫バス_ERR:
    pERR_MES = pPROGRAM_ID + "/ MDA010_長短振替表_神姫バス() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub


