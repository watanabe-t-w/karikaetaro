Attribute VB_Name = "MAA030_銀行マスタ"
Option Explicit
'
Private Const pPROGRAM_ID As String = "MAA030_銀行マスタ"

Dim w銀行マスタ() As MAA030_銀行
'------------------------< 銀行マスタ >-------------------------
Type MAA030_銀行
    銀行番号 As String
    銀行名 As String
    
    支払日 As Integer
    営業日 As String
    利息区分 As String
    利息日数 As String
    利息支払 As String
    利息控除 As String
    金利計算 As String
    取消フラグ As Integer
End Type
'
'------------------------------------------------
' MAA030_銀行マスタ設定
'------------------------------------------------
Public Sub MAA030_銀行マスタ設定()
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    Dim j As Integer
'
    On Error GoTo MAA030_銀行マスタ設定_ERR
'
    wstr = ""
    wstr = wstr + "Select Count(*) As カウント From DAAA040_銀行マスタ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If wRs("カウント") = 0 Then
            wRs.Close
            Set wRs = Nothing
            
            Exit Sub
        End If
    
        ReDim w銀行マスタ(wRs("カウント") - 1)
    wRs.Close
    Set wRs = Nothing
'
    wstr = ""
    wstr = wstr + "Select * From DAAA040_銀行マスタ"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        j = -1
        
        Do Until wRs.EOF
            j = j + 1
                
            w銀行マスタ(j).銀行番号 = wRs("銀行番号")
            w銀行マスタ(j).銀行名 = wRs("銀行名")

            w銀行マスタ(j).支払日 = wRs("支払日")
            w銀行マスタ(j).営業日 = wRs("営業日区分")
            w銀行マスタ(j).利息区分 = wRs("利息区分")
            w銀行マスタ(j).利息日数 = wRs("利息計算日数区分")
            w銀行マスタ(j).利息支払 = wRs("利息支払方法")
            w銀行マスタ(j).利息控除 = wRs("利息控除区分")
            w銀行マスタ(j).金利計算 = wRs("金利計算年間日数")
            w銀行マスタ(j).取消フラグ = wRs("取消フラグ")
             
            wRs.MoveNext
        Loop
    
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA030_銀行マスタ設定_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA030_銀行マスタ設定() でエラー" + vbCrLf + vbCrLf + _
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
' MAA030_銀行マスタRead
'------------------------------------------------
Public Function MAA030_銀行マスタRead(p銀行番号 As String) As MAA030_銀行
'
    Dim j As Integer
'
    On Error GoTo MAA030_銀行マスタRead_ERR
'
    MAA030_銀行マスタRead.銀行番号 = ""
    MAA030_銀行マスタRead.銀行名 = ""
    
    MAA030_銀行マスタRead.支払日 = 0
    MAA030_銀行マスタRead.営業日 = ""
    MAA030_銀行マスタRead.利息区分 = ""
    MAA030_銀行マスタRead.利息日数 = ""
    MAA030_銀行マスタRead.利息支払 = ""
    MAA030_銀行マスタRead.利息控除 = ""
    MAA030_銀行マスタRead.金利計算 = ""
    MAA030_銀行マスタRead.取消フラグ = 0

    For j = 0 To UBound(w銀行マスタ)
        If UCase(w銀行マスタ(j).銀行番号) = UCase(p銀行番号) Then
            
            MAA030_銀行マスタRead.銀行番号 = w銀行マスタ(j).銀行番号
            MAA030_銀行マスタRead.銀行名 = w銀行マスタ(j).銀行名

            MAA030_銀行マスタRead.支払日 = w銀行マスタ(j).支払日
            MAA030_銀行マスタRead.営業日 = w銀行マスタ(j).営業日
            MAA030_銀行マスタRead.利息区分 = w銀行マスタ(j).利息区分
            MAA030_銀行マスタRead.利息日数 = w銀行マスタ(j).利息日数
            MAA030_銀行マスタRead.利息支払 = w銀行マスタ(j).利息支払
            MAA030_銀行マスタRead.利息控除 = w銀行マスタ(j).利息控除
            MAA030_銀行マスタRead.金利計算 = w銀行マスタ(j).金利計算
            MAA030_銀行マスタRead.取消フラグ = w銀行マスタ(j).取消フラグ
            
            Exit For
        End If
    Next
'
    Exit Function
'
'----------< ERROR ROUTINE >---------------------------------------------------
MAA030_銀行マスタRead_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA030_銀行マスタRead() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function
