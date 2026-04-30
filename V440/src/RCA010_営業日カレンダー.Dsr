VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RCA010_営業日カレンダー 
   Caption         =   "営業日カレンダー"
   ClientHeight    =   8115
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14910
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   _ExtentX        =   26300
   _ExtentY        =   14314
   SectionData     =   "RCA010_営業日カレンダー.dsx":0000
End
Attribute VB_Name = "RCA010_営業日カレンダー"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "RCA010_営業カレンダー"

Private Cweekday(11) As MCL010_Typeレコード
Private Type MCL010_Typeレコード       '縦6横7 42コのセル
    xDate(42)  As String
    xHyoji(42)  As String      'セルに該当日がない場合は空白
    xKyujitu(42) As Boolean    '休日フラグ(赤)
End Type
'
'------------------------------------------------
' ActiveReport_DataInitialize
'------------------------------------------------
Private Sub ActiveReport_DataInitialize()
    Dim M As Integer
    Dim l As Integer
    Dim wi01 As Integer
    Dim wv01 As Variant
    Dim ws01 As String
'
    For M = 1 To 12
        wi01 = 0
        For l = 1 To 31
            wv01 = GRpt.テキスト_01 & "/" & Right$("00" & M, 2) & "/" & Right$("00" & l, 2)
            If Not IsDate(wv01) Then
                Exit For
            End If
    
            If l = 1 Then
                wi01 = Weekday(CDate(wv01)) - 1 'vbSunday(1),vbSaturday(7)
            End If
    
            Call C休日.計算(CDate(wv01), 0)
            GRet = C休日.休日
    
            Cweekday(M - 1).xDate(l + wi01) = GRpt.テキスト_01 & Right$("00" & M, 2) & Right$("00" & l, 2)
            Cweekday(M - 1).xHyoji(l + wi01) = CStr(l)
            Cweekday(M - 1).xKyujitu(l + wi01) = GRet
        Next l
    Next M
'
    For M = 1 To 12
        Me.Detail.Controls("T_M" & CStr(M)).Text = CStr(M) & " 月"
        For l = 1 To 42
            ws01 = "Field" & CStr(l + ((M - 1) * 42))
            Me.Detail.Controls(ws01).Text = Cweekday(M - 1).xHyoji(l)
            If Cweekday(M - 1).xKyujitu(l) = True Then
                Me.Detail.Controls(ws01).ForeColor = C_Red
            Else
                Me.Detail.Controls(ws01).ForeColor = C_Black
            End If
        Next l
    Next M
'
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
    
    Me.PageHeader.Controls("L_帳票名").Caption = GRpt.テキスト_01 & "年 営業日カレンダー"
'
End Sub
