Attribute VB_Name = "MAA020_コントロール"
Option Explicit
'
Private Const pPROGRAM_ID As String = "MAA020_コントロール"

'------------------------< コントロール ファイル >-------------------------
Type MAA020_コントロールファイル
    最終実績年月 As Date
'    サーバーフォルダ As String
End Type
'
'------------------------------------------------
' MAA020_コントロールファイル_Read
'------------------------------------------------
Public Sub MAA020_コントロールファイル_Read()
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
'
    On Error GoTo MAA020_コントロールファイル_Read_ERR
'
    wstr = ""
    wstr = wstr + "Select * From DAAA020_コントロール"
    wstr = wstr + " Where System = 'System'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If Not wRs.EOF Then
            Gコントロール.最終実績年月 = wRs("最終実績年月")
'            Gコントロール.サーバーフォルダ = P8.FCStr(wRs("サーバーフォルダ"))
        End If
    wRs.Close
    Set wRs = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >--------------------------------------------------------
MAA020_コントロールファイル_Read_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA020_コントロールファイル_Read() でエラー" + vbCrLf + vbCrLf + _
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
' MAA020_最終実績LIST
'------------------------------------------------
Public Sub MAA020_最終実績LIST(pKeyName As String)
'
    Dim wLDb As New ADODB.Connection
    Dim wRs As ADODB.Recordset
    Dim wstr As String
'
    On Error GoTo MAA020_最終実績LIST_ERR
'
    '----------< List.mdb Open >----------------------------------------------------
    Call AdoDbOpen("Jet", wLDb, GSerDir + "\" + GMain, "", , GPwd)
    
        wstr = ""
        wstr = wstr + "Select 最終実績年月"
        wstr = wstr + " From DAAA070_企業名マスタ"
        wstr = wstr + " Where 企業名Key='" + pKeyName + "'"
        Call AdoRecordsetOpen(wLDb, wRs, wstr)
        
            wRs("最終実績年月") = Format(Gコントロール.最終実績年月, "yyyy/mm/dd")
            wRs.Update
            
        wRs.Close
        Set wRs = Nothing
        
    wLDb.Close
    Set wLDb = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >--------------------------------------------------------
MAA020_最終実績LIST_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA020_最終実績LIST() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub


