Attribute VB_Name = "MAA090_祝日マスタ"
Option Explicit
'
Private Const pPROGRAM_ID As String = "MAA090_祝日マスタ"

'------------------------< 祝日マスタ >-------------------------
Type MAA090_祝日
    年月日 As Variant
    名称 As String
End Type

Type MAA090_KakuninMsg
    借入番号 As String
    銀行番号 As String
    銀行名 As String
    確認年月日 As String
End Type

'------------------------------------------------
' MAA090_祝日マスタ設定
'------------------------------------------------
Public Sub MAA090_祝日マスタ設定()
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    Dim j As Integer
'
    On Error GoTo MAA090_祝日マスタ設定_ERR
'
    wstr = ""
    wstr = wstr + "Select Count(*) As カウント From DACA010_祝日マスタ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs("カウント") = 0 Then
            wRs.Close
            Set wRs = Nothing
            
            ReDim G祝日マスタ(0)
            
            Exit Sub
        End If
    
        ReDim G祝日マスタ(wRs("カウント") - 1)
    wRs.Close
    Set wRs = Nothing
'
    wstr = ""
    wstr = wstr + "Select * From DACA010_祝日マスタ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        j = -1
        
        Do Until wRs.eof
            j = j + 1
                
            G祝日マスタ(j).年月日 = wRs("年月日")
            G祝日マスタ(j).名称 = wRs("名称")
             
            wRs.MoveNext
        Loop
    
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA090_祝日マスタ設定_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA090_祝日マスタ設定() でエラー" + vbCrLf + vbCrLf + _
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
' MAA090_祝日マスタRead
'------------------------------------------------
Public Function MAA090_祝日マスタRead(p年月日 As Variant) As MAA090_祝日
'
    Dim j As Integer
'
    On Error GoTo MAA090_祝日マスタRead_ERR
'
    MAA090_祝日マスタRead.年月日 = ""
    MAA090_祝日マスタRead.名称 = ""

    For j = 0 To UBound(G祝日マスタ)
        If UCase(G祝日マスタ(j).年月日) = UCase(p年月日) Then
            
            MAA090_祝日マスタRead.年月日 = G祝日マスタ(j).年月日
            MAA090_祝日マスタRead.名称 = G祝日マスタ(j).名称
            
            Exit For
        End If
    Next
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA090_祝日マスタRead_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA090_祝日マスタRead() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function


