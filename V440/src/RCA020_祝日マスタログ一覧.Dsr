VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RCA020_祝日マスタログ一覧 
   Caption         =   "祝日マスタログ一覧"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13725
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   _ExtentX        =   24209
   _ExtentY        =   10239
   SectionData     =   "RCA020_祝日マスタログ一覧.dsx":0000
End
Attribute VB_Name = "RCA020_祝日マスタログ一覧"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "RCA020_祝日マスタログ一覧"

Public bLastIsSingle As Boolean
Private iRow As Integer

'------------------------------------------------
' ActiveReport_DataInitialize
'------------------------------------------------
Private Sub ActiveReport_DataInitialize()
    Fields.Add "I_借入番号"
    Fields.Add "I_銀行番号"
    Fields.Add "I_銀行名"
    Fields.Add "I_確認年月日"
    iRow = LBound(GCal) + 1
End Sub

'------------------------------------------------
' ActiveReport_FetchData
'------------------------------------------------
Private Sub ActiveReport_FetchData(eof As Boolean)
    If iRow > UBound(GCal) Then
        eof = True
        Exit Sub
    End If
    Fields("I_借入番号") = GCal(iRow).借入番号
    Fields("I_銀行番号") = GCal(iRow).銀行番号
    Fields("I_銀行名") = GCal(iRow).銀行名
    Fields("I_確認年月日") = GCal(iRow).確認年月日

    ' 複数のレコードがある場合、eofをFalseにすることが重要です。
    eof = False
    iRow = iRow + 1
End Sub

'------------------------------------------------
' ActiveReport_ReportStart
'------------------------------------------------
Private Sub ActiveReport_ReportStart()
    '----------------------------------------------------------------
    '                           ** 見出し **
    '----------------------------------------------------------------
    出力日 = Now
    企業名 = GCoName
End Sub

'------------------------------------------------
' ActiveReport_ReportEnd
'------------------------------------------------
Private Sub ActiveReport_ReportEnd()
    Call MX040_CsvOut_KARI
End Sub

