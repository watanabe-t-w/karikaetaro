Attribute VB_Name = "MAA060_税率マスタ"
Option Explicit
'
Private Const pPROGRAM_ID As String = "MAA060_税率マスタ"

'------------------------< 税率マスタ >-------------------------
Type MAA060_税率
    年月 As Date
    不課税率 As Double
    課税率 As Double
    非課税率 As Double
    法人税率 As Double
End Type

'------------------------------------------------
' MAA060_税率マスタ設定
'------------------------------------------------
Public Sub MAA060_税率マスタ設定()
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    Dim j As Integer
'
    On Error GoTo MAA060_税率マスタ設定_ERR
'
    wstr = ""
    wstr = wstr + "Select Count(*) As カウント From DAAA060_税率マスタ"
    wstr = wstr + " Where 取消フラグ=0"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs("カウント") = 0 Then
            wRs.Close
            Set wRs = Nothing
            
            Exit Sub
        End If
    
        ReDim G税率マスタ(wRs("カウント") - 1)
        
    wRs.Close
    Set wRs = Nothing
'
    wstr = ""
    wstr = wstr + "Select * From DAAA060_税率マスタ"
    wstr = wstr + " Where 取消フラグ=0"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        j = -1
        
        Do Until wRs.EOF
            j = j + 1
                
            G税率マスタ(j).年月 = wRs("年月")
            G税率マスタ(j).不課税率 = wRs("不課税率")
            G税率マスタ(j).課税率 = wRs("課税率")
            G税率マスタ(j).非課税率 = wRs("非課税率")
            G税率マスタ(j).法人税率 = wRs("法人税率")
             
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA060_税率マスタ設定_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA060_税率マスタ設定() でエラー" + vbCrLf + vbCrLf + _
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
' MAA060_税率マスタRead
'------------------------------------------------
Public Function MAA060_税率マスタRead(p年月 As Date) As MAA060_税率
'
    Dim j As Integer
    Dim wdate As Date
'
    On Error GoTo MAA060_税率マスタRead_ERR
'
    p年月 = Format(p年月, "yyyy/mm/dd")
    
    MAA060_税率マスタRead.不課税率 = 0
    MAA060_税率マスタRead.課税率 = 0
    MAA060_税率マスタRead.非課税率 = 0
    MAA060_税率マスタRead.法人税率 = 0
'
    For j = 0 To UBound(G税率マスタ)
        wdate = Format(G税率マスタ(j).年月, "yyyy/mm/dd")
        If j = 0 Then
            MAA060_税率マスタRead.年月 = G税率マスタ(j).年月
            MAA060_税率マスタRead.不課税率 = G税率マスタ(j).不課税率
            MAA060_税率マスタRead.課税率 = G税率マスタ(j).課税率
            MAA060_税率マスタRead.非課税率 = G税率マスタ(j).非課税率
            MAA060_税率マスタRead.法人税率 = G税率マスタ(j).法人税率
            If wdate >= p年月 Then
                Exit For
            End If
        Else
            If wdate > p年月 Then
                Exit For
            ElseIf wdate = p年月 Then
                MAA060_税率マスタRead.年月 = G税率マスタ(j).年月
                MAA060_税率マスタRead.不課税率 = G税率マスタ(j).不課税率
                MAA060_税率マスタRead.課税率 = G税率マスタ(j).課税率
                MAA060_税率マスタRead.非課税率 = G税率マスタ(j).非課税率
                MAA060_税率マスタRead.法人税率 = G税率マスタ(j).法人税率

                Exit For
            End If
            MAA060_税率マスタRead.年月 = G税率マスタ(j).年月
            MAA060_税率マスタRead.不課税率 = G税率マスタ(j).不課税率
            MAA060_税率マスタRead.課税率 = G税率マスタ(j).課税率
            MAA060_税率マスタRead.非課税率 = G税率マスタ(j).非課税率
            MAA060_税率マスタRead.法人税率 = G税率マスタ(j).法人税率
        End If
    Next
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA060_税率マスタRead_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA060_税率マスタRead() でエラー" + vbCrLf + vbCrLf + _
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
' MAA060_指定年月課税率
'------------------------------------------------
Public Function MAA060_指定年月課税率(p年月 As Date) As Double
'
    Dim j As Integer
    Dim wdate As Date
    Dim wd01 As Double
'
    On Error GoTo MAA060_指定年月課税率_ERR
'
    p年月 = Format(p年月, "yyyy/mm/dd")
    wd01 = 0
    
    For j = 0 To UBound(G税率マスタ)
        wdate = Format(G税率マスタ(j).年月, "yyyy/mm/dd")
        
        If j = 0 Then
            wd01 = G税率マスタ(j).課税率
            If wdate >= p年月 Then
                Exit For
            End If
        Else
            If wdate > p年月 Then
                Exit For
            ElseIf wdate = p年月 Then
                wd01 = G税率マスタ(j).課税率
                Exit For
            End If
            wd01 = G税率マスタ(j).課税率
        End If
    Next
'
    MAA060_指定年月課税率 = wd01
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA060_指定年月課税率_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA060_指定年月課税率() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function
