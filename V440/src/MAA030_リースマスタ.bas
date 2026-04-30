Attribute VB_Name = "MAA030_リースマスタ"
Option Explicit
'
Private Const pPROGRAM_ID As String = "MAA030_リースマスタ"
'
Dim wリースマスタ() As MAA030_リース
'------------------------< リースマスタ >-------------------------
Type MAA030_リース
    リース会社番号 As String
    リース会社名 As String
    支払日 As Integer
    営業日区分 As String
    取消フラグ As Integer
End Type
'
'------------------------------------------------
' MAA030_リースマスタ設定
'------------------------------------------------
Public Sub MAA030_リースマスタ設定()
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    Dim j As Integer
'
    On Error GoTo MAA030_リースマスタ設定_ERR
'
    wstr = ""
    wstr = wstr + "Select Count(*) As カウント From DAAA040_リースマスタ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs("カウント") = 0 Then
            wRs.Close
            Set wRs = Nothing
            
            Exit Sub
        End If
    
        ReDim wリースマスタ(wRs("カウント") - 1)
    wRs.Close
    Set wRs = Nothing
'
    wstr = ""
    wstr = wstr + "Select * From DAAA040_リースマスタ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        j = -1
        
        Do Until wRs.EOF
            j = j + 1
                
            wリースマスタ(j).リース会社番号 = wRs("リース会社番号")
            wリースマスタ(j).リース会社名 = wRs("リース会社名")
            wリースマスタ(j).支払日 = wRs("支払日")
            wリースマスタ(j).営業日区分 = wRs("営業日区分")
            wリースマスタ(j).取消フラグ = wRs("取消フラグ")
             
            wRs.MoveNext
        Loop
    
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA030_リースマスタ設定_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA030_リースマスタ設定() でエラー" + vbCrLf + vbCrLf + _
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
' MAA030_リースマスタRead
'------------------------------------------------
Public Function MAA030_リースマスタRead(pリース会社番号 As String) As MAA030_リース
'
    Dim j As Integer
'
    On Error GoTo MAA030_リースマスタRead_ERR
'
    MAA030_リースマスタRead.リース会社番号 = ""
    MAA030_リースマスタRead.リース会社名 = ""
    MAA030_リースマスタRead.支払日 = 0
    MAA030_リースマスタRead.営業日区分 = ""
    MAA030_リースマスタRead.取消フラグ = 0

    For j = 0 To UBound(wリースマスタ)
        If UCase(wリースマスタ(j).リース会社番号) = UCase(pリース会社番号) Then
            
            MAA030_リースマスタRead.リース会社番号 = wリースマスタ(j).リース会社番号
            MAA030_リースマスタRead.リース会社名 = wリースマスタ(j).リース会社名
            MAA030_リースマスタRead.支払日 = wリースマスタ(j).支払日
            MAA030_リースマスタRead.営業日区分 = wリースマスタ(j).営業日区分
            MAA030_リースマスタRead.取消フラグ = wリースマスタ(j).取消フラグ
            
            Exit For
        End If
    Next
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA030_リースマスタRead_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA030_リースマスタRead() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function
