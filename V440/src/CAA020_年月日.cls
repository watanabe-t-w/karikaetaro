VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAA020_年月日"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'********************************************
Public Function CharCheck(pdata As String) As String
'********************************************
    Dim wInt As Integer
    Dim j As Integer
    Dim wstr As String: wstr = ""
'
    If IsNull(pdata) Or pdata = "" Then
        CharCheck = ""
        Exit Function
    End If
    
    wInt = Len(pdata)
    For j = 1 To wInt
        If IsNumeric(Mid(pdata, j, 1)) Then
            wstr = wstr + Mid(pdata, j, 1)
        End If
    Next
'
    CharCheck = wstr
'
End Function

'********************************************
Public Function GetDate(pType As String, p年月日 As Variant, Optional p日 As Integer = 0) As Variant
'********************************************
    Dim wdate As Date
    Dim w年 As Integer
    Dim w月 As Integer
    Dim w閏年 As Integer  'V180 07/01/20
    Dim w日 As Integer    'V180 07/01/21
'
    If IsNull(p年月日) Then
        GetDate = Null
        Exit Function
    End If

    w年 = Year(p年月日)
    w月 = Month(p年月日)
    w日 = p日               'V180 07/01/20
    w閏年 = w年 Mod 4       'V180 07/01/20
    If (p日 = 30 Or p日 = 29) And w月 = 2 Then             'V180 07/01/20
        If w閏年 = 0 Then                   'V180 07/01/20
            w日 = 29                        'V180 07/01/20
        Else                                'V180 07/01/20
            w日 = 28                        'V180 07/01/20
        End If                              'V180 07/01/20
    End If                                  'V180 07/01/20
    
            
    wdate = DateSerial(w年, w月, 1)

    Select Case pType
        Case "月始"
            GetDate = wdate
        Case "月末"
            wdate = DateAdd("m", 1, wdate)
            wdate = DateAdd("d", -1, wdate)
            GetDate = wdate
        Case "設定"
            If w日 = 0 Then
                GetDate = wdate
            ElseIf w日 = 31 Then
                wdate = DateAdd("m", 1, wdate)
                wdate = DateAdd("d", -1, wdate)
                GetDate = wdate
            Else
                GetDate = DateSerial(w年, w月, w日)
            End If
    End Select
'
End Function

'********************************************
Public Function FormatDate(pType As String, p年月日 As String) As String
'********************************************
'
    Dim wstr As String
    Dim w年月日 As String
    
    w年月日 = p年月日
'
    p年月日 = CharCheck(p年月日)
    
    '2011.11.17 By m.mino
'    p年月日 = Replace(p年月日, "/", "")
    
    If Not IsNumeric(p年月日) Then
        FormatDate = p年月日
        Exit Function
    End If
       
    wstr = LCase(p年月日)
    If Len(wstr) = 1 _
    Or Len(wstr) = 3 _
    Or Len(wstr) = 5 Then
        wstr = "0" + wstr
    End If
    
    If G基本情報.日付入力区分 = "0" Then
        p年月日 = Replace(p年月日, "/", "")
        Select Case pType
            Case "年", "月", "日"
                FormatDate = Left(wstr, 2) + pType
            Case "年月"
                FormatDate = Left(wstr, 2) + "年" + Mid(wstr, 3, 2) + "月"
            Case "月日"
                FormatDate = Left(wstr, 2) + "月" + Mid(wstr, 3, 2) + "日"
            Case "年月日"
                FormatDate = Left(wstr, 2) + "年" + Mid(wstr, 3, 2) + "月" + Mid(wstr, 5, 2) + "日"
        End Select
    Else
        FormatDate = FormatDateSla(pType, w年月日)
    End If
'
End Function

'********************************************
Public Function 平成To西暦(pType As String, p年月日 As String, Optional pNullOk As Boolean = False) As Variant
'********************************************
    Dim wstr As String
'
    If pNullOk _
    And p年月日 = "" Then
        平成To西暦 = Null
        Exit Function
    End If

    wstr = p年月日
    If pType = "年" Then
        wstr = wstr + "年01月"
    End If
    
    If G基本情報.日付入力区分 = 0 Then
        wstr = "平成" + wstr
    Else
    End If
    If Not IsDate(wstr) Then
        '** エラー **
        平成To西暦 = 0
        Exit Function
    End If
     
    平成To西暦 = CDate(wstr)

End Function

'********************************************
Public Function 年度開始年月日(p年 As String, p表示 As String) As String
'********************************************
    Dim wstr As String
    Dim w年月日 As Date
'
    If Not IsNumeric(p年) Then
        年度開始年月日 = ""
        Exit Function
    End If
    
    Select Case Len(p年)
    Case 4
        wstr = p年 + "/" + CStr(G基本情報.決算月)
        If Not IsDate(wstr) Then
            年度開始年月日 = ""
            Exit Function
        End If
        w年月日 = DateAdd("m", -11, CDate(wstr))
        w年月日 = CDate(p年 + "/" + Format(w年月日, "mm/dd"))
    Case 1, 2
        wstr = "平成" + p年 + "年" + CStr(G基本情報.決算月) + "月"
        If Not IsDate(wstr) Then
            年度開始年月日 = ""
            Exit Function
        End If
        w年月日 = DateAdd("m", -11, CDate(wstr))
        w年月日 = CDate("平成" + p年 + "年" + Format(w年月日, "mm月dd日"))
    End Select
    
    If p表示 = "西暦" Then
        年度開始年月日 = w年月日
    ElseIf p表示 = "平成" Then
        年度開始年月日 = Format(w年月日, "gggee年mm月dd日")
    ElseIf p表示 = "y" Then
        年度開始年月日 = Format(w年月日, "yyyy")
    ElseIf p表示 = "e" Then
        年度開始年月日 = Format(w年月日, "ee")
    End If
'
End Function

'********************************************
Public Function 年度開始年月日_P決算月(p年 As String, p表示 As String, p決算月 As String) As String
'********************************************
    Dim wstr As String
    Dim w年月日 As Date
'
    If Not IsNumeric(p年) Then
        年度開始年月日_P決算月 = ""
        Exit Function
    End If
    
    Select Case Len(p年)
    Case 4
        wstr = p年 + "/" + p決算月
        If Not IsDate(wstr) Then
            年度開始年月日_P決算月 = ""
            Exit Function
        End If
        w年月日 = DateAdd("m", -11, CDate(wstr))
        w年月日 = CDate(p年 + "/" + Format(w年月日, "mm/dd"))
    Case 1, 2
        wstr = "平成" + p年 + "年" + p決算月 + "月"
        If Not IsDate(wstr) Then
            年度開始年月日_P決算月 = ""
            Exit Function
        End If
        w年月日 = DateAdd("m", -11, CDate(wstr))
        w年月日 = CDate("平成" + p年 + "年" + Format(w年月日, "mm月dd日"))
    End Select
    
    If p表示 = "西暦" Then
        年度開始年月日_P決算月 = w年月日
    ElseIf p表示 = "平成" Then
        年度開始年月日_P決算月 = Format(w年月日, "gggee年mm月dd日")
    ElseIf p表示 = "y" Then
        年度開始年月日_P決算月 = Format(w年月日, "yyyy")
    ElseIf p表示 = "e" Then
        年度開始年月日_P決算月 = Format(w年月日, "ee")
    End If
'
End Function

'********************************************
Public Function 指定開始年月日(p年 As String, p表示 As String, p決算月 As String) As String
'********************************************
    Dim wstr As String
    Dim w年月日 As Date
'
    If Not IsNumeric(p年) Then
        指定開始年月日 = ""
        Exit Function
    End If
    
    Select Case Len(p年)
    Case 4
        wstr = p年 + "/" + p決算月
        If Not IsDate(wstr) Then
            指定開始年月日 = ""
            Exit Function
        End If
        w年月日 = DateAdd("m", -11, CDate(wstr))
        w年月日 = CDate(p年 + "/" + Format(w年月日, "mm/dd"))
    Case 1, 2
        wstr = "平成" + p年 + "年" + p決算月 + "月"
        If Not IsDate(wstr) Then
            指定開始年月日 = ""
            Exit Function
        End If
        w年月日 = DateAdd("m", -11, CDate(wstr))
        w年月日 = CDate("平成" + p年 + "年" + Format(w年月日, "mm月dd日"))
    End Select
    
    If p表示 = "西暦" Then
        指定開始年月日 = w年月日
    ElseIf p表示 = "平成" Then
        指定開始年月日 = Format(w年月日, "gggee年mm月dd日")
    ElseIf p表示 = "y" Then
        指定開始年月日 = Format(w年月日, "yyyy")
    ElseIf p表示 = "e" Then
        指定開始年月日 = Format(w年月日, "ee")
    End If
'
End Function

'********************************************
Public Function 年度変換(pDate As String) As Integer
'********************************************
'
    Dim wstr As String
    Dim w年月日 As Date, 本年度 As Date, 来年度 As Date
'
    If Not IsDate(pDate) Then
        年度変換 = 0
        Exit Function
    End If
    w年月日 = CDate(pDate)
    
    wstr = Format(pDate, "yy") + "/" + Right("00" + CStr(G基本情報.決算月), 2) + "/01"
    If Not IsDate(wstr) Then
        年度変換 = 0
        Exit Function
    End If
    本年度 = CDate(wstr)
    本年度 = DateAdd("m", -11, CDate(wstr))
    本年度 = Format(pDate, "yy") + "/" + Format(本年度, "mm/dd")
    If w年月日 < 本年度 Then
        本年度 = DateAdd("yyyy", -1, 本年度)
    End If
    
    来年度 = DateAdd("yyyy", 1, 本年度)
    
    If 本年度 <= w年月日 And w年月日 < 来年度 Then
        年度変換 = Format(本年度, "yyyy")
    Else
        年度変換 = Format(w年月日, "yyyy")
    End If
'
End Function
'********************************************
Public Function FormatDateSla(pType As String, p年月日 As String)
'********************************************

    Dim w年 As String
    Dim w月 As String
    Dim w日 As String
    
    Dim wBuf() As String
    
    If InStr(p年月日, "/") = 0 Then
        Select Case pType
            Case "年"
                FormatDateSla = Left(p年月日, 4)
            Case "月", "日"
                FormatDateSla = Left(p年月日, 2)
            Case "年月"
                FormatDateSla = Left(p年月日, 4) + "/" + Format(Mid(p年月日, 5, 2), "00") + ""
            Case "月日"
                FormatDateSla = Left(p年月日, 2) + "/" + Format(Mid(p年月日, 3, 2), "00") + ""
            Case "年月日"
                If Mid(p年月日, 5, 2) > 12 Then
                    w年 = Left(p年月日, 4)
                    w月 = Format(Mid(p年月日, 5, 1), "00")
                    w日 = Format(Mid(p年月日, 6, 2), "00")
                Else
                    w年 = Left(p年月日, 4)
                    w月 = Format(Mid(p年月日, 5, 2), "00")
                    w日 = Format(Mid(p年月日, 7, 2), "00")
                End If
                
                If w月 = "" Then
                    w月 = "01"
                End If
                
                If w日 = "" Then
                    w日 = "01"
                End If
                
                FormatDateSla = w年 & "/" & w月 & "/" & w日
                
        End Select
    Else
        wBuf = Split(p年月日, "/")
        If UBound(wBuf) = 0 Then
            w年 = wBuf(0)
        ElseIf UBound(wBuf) = 1 Then
            w年 = wBuf(0)
            w月 = Format(wBuf(1), "00")

        ElseIf UBound(wBuf) = 2 Then
            w年 = wBuf(0)
            w月 = Format(wBuf(1), "00")
            w日 = Format(wBuf(2), "00")

        End If
        
        If w月 = "" Then
            w月 = "01"
        End If
        
        If w日 = "" Then
            w日 = "01"
        End If
        
        Select Case pType
            Case "年"
                p年月日 = Replace(p年月日, "/", "")
                FormatDateSla = Left(p年月日, 4)
                FormatDateSla = w年
            Case "月", "日"
                p年月日 = Replace(p年月日, "/", "")
                FormatDateSla = Left(p年月日, 2)
            Case "年月"
                p年月日 = Replace(p年月日, "/", "")
                FormatDateSla = Left(p年月日, 4) + "/" + Format(Mid(p年月日, 5, 2), "00") + ""
            Case "月日"
                p年月日 = Replace(p年月日, "/", "")
                FormatDateSla = Left(p年月日, 2) + "/" + Format(Mid(p年月日, 3, 2), "00") + ""
            Case "年月日"
                FormatDateSla = w年 & "/" & w月 & "/" & w日
        End Select
        
    End If
    

    
End Function

'********************************************
Public Function 前期末決算日算出(p年月日 As String) As Variant
'********************************************
'
    Dim wstr As String
    Dim w年月日 As Date
    Dim wv01 As Variant
'
    If p年月日 = "" Then
        Exit Function
    End If
    
    wv01 = 年度開始年月日(Format(p年月日, "yyyy"), "西暦")
    wv01 = DateAdd("d", -1, P8.FCDate(wv01))
    If P8.FCDate(wv01) >= P8.FCDate(p年月日) Then
        wv01 = DateAdd("yyyy", -1, wv01)
    End If
'
    前期末決算日算出 = wv01
'
End Function


