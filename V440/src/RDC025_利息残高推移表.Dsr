VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RDC025_利息残高推移表 
   Caption         =   "利息残高推移表"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13455
   Icon            =   "RDC025_利息残高推移表.dsx":0000
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   _ExtentX        =   23733
   _ExtentY        =   14684
   SectionData     =   "RDC025_利息残高推移表.dsx":0ECA
End
Attribute VB_Name = "RDC025_利息残高推移表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "RDC025_利息残高推移表"
'
Dim ws_Risoku As String, ws_Tyotan As String
Dim wsTL(6) As String

Dim wd_Saki(4, 12, 4) As Double, wd_Ato(4, 12, 4) As Double 'GRセクション数,列数,Field数
Dim Cnt_Saki(4) As Double, Cnt_Ato(4) As Double 'GRセクション数

Dim wML As Integer, wFD As Integer
Dim w番号 As String, wsTbl As String, wsTbl2 As String
Dim w分母 As Integer
Dim w推移表タイトル As MAA910_推移表タイトル
'
'------------------------------------------------
' ActiveReport_ReportStart
'------------------------------------------------
Private Sub ActiveReport_ReportStart()
'
    Dim wRs As ADODB.Recordset
    Dim wWhere As String, wOrder As String
    Dim FLG_Order As Boolean
    
    Dim wstr As String
    Dim wsRet As String
    Dim ws_Ginko As String
    Dim j As Integer, k As Integer, l As Integer
    
    Dim wdate As Date
    Dim w開始年月日 As Date
    Dim w推移表区分 As String
    Dim wsS As String, ws01 As String
'
    On Error GoTo ActiveReport_ReportStart_ERR
'
    '----------------------------------------------------------------
    '                         ** 初期設定 **
    '----------------------------------------------------------------
    'Connection
    Me.DataControl1.Connection = GDb
   
    '用紙セット
    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = ddOLandscape

    wML = 12    '列数
    wFD = 5     'Field数
'
    'パラメータ部分
    'Me.PageHeader.Controls("L_計画番号").Visible = True
    'Me.PageHeader.Controls("L_最終実績年月").Visible = True
    Me.PageHeader.Controls("L_金融リストラ番号").Visible = True
    
    'Me.PageHeader.Controls("H00_借入計画番号").Visible = True
    'Me.PageHeader.Controls("H00_最終実績年月").Visible = True
    Me.PageHeader.Controls("H00_金融リストラ番号").Visible = True
    
    'Me.PageHeader.Controls("L_金融リストラ番号").Left = 4535
    'Me.PageHeader.Controls("H00_金融リストラ番号").Left = 4535
    
    'Me.PageHeader.Controls("Line5").Visible = True
    'Me.PageHeader.Controls("Line6").Visible = True
    
    'Me.PageHeader.Controls("Shape1").Width = 6803
    'Me.PageHeader.Controls("Shape2").Width = 6803
    'Me.PageHeader.Controls("Shape3").Width = 6803
    'Me.PageHeader.Controls("Line4").X2 = 6803
    
    'Me.PageHeader.Controls("L_借入計画番号").Visible = True
    'Me.Detail.Controls("I_借入計画番号").Visible = True
    
    If GProduct <> "金剛石" Then
        '帳票名
        'Me.PageHeader.Controls("L_帳票名").Left = 4819
        
        'パラメータ部分
        'Me.PageHeader.Controls("L_計画番号").Visible = False
        'Me.PageHeader.Controls("L_最終実績年月").Visible = False
        
        'Me.PageHeader.Controls("H00_借入計画番号").Visible = False
        'Me.PageHeader.Controls("H00_最終実績年月").Visible = False
        
        'Me.PageHeader.Controls("Line5").Visible = False
        'Me.PageHeader.Controls("Line6").Visible = False
        
        'If GRpt.金融 = "" Then
        '    Me.PageHeader.Controls("L_金融リストラ番号").Visible = False
        '    Me.PageHeader.Controls("H00_金融リストラ番号").Visible = False
        '
        '    Me.PageHeader.Controls("Shape1").Visible = False
        '    Me.PageHeader.Controls("Shape2").Visible = False
        '    Me.PageHeader.Controls("Shape3").Visible = False
        '    Me.PageHeader.Controls("Line4").Visible = False
        'Else
        '    Me.PageHeader.Controls("L_金融リストラ番号").Left = 0
        '    Me.PageHeader.Controls("H00_金融リストラ番号").Left = 0
        '
        '    Me.PageHeader.Controls("Shape1").Width = 2268
        '    Me.PageHeader.Controls("Shape2").Width = 2268
        '    Me.PageHeader.Controls("Shape3").Width = 2268
        '    Me.PageHeader.Controls("Line4").X2 = 2268
        'End If
    
        'Me.PageHeader.Controls("L_借入計画番号").Visible = False
        'Me.Detail.Controls("I_借入計画番号").Visible = False
    End If
'
    '貸借設定、wsTbl設定
    wsTbl = "": wsTbl2 = ""
    ws_Ginko = ""
    
    Select Case GRpt.帳票名
    Case "借入利息残高推移表", "借入利息残高推移表 プロジェクト"
        If GRpt.選択 = "分岐点借入金" Then
            wsTbl = "DBDA010_分岐点借入金"
            wsTbl2 = "DBDA010_借入金"
        Else
            wsTbl = "DBDA010_借入金"
        End If

        ws_Ginko = "外部借入"

        Me.PageHeader.Controls("L_借入番号").Caption = "借入番号"
        'Me.PageHeader.Controls("L_借入内容").Caption = "借入内容"
        'Me.PageHeader.Controls("L_計画番号").Caption = "借入計画番号"

    Case "貸付利息残高推移表", "貸付利息残高推移表 プロジェクト"
        If GRpt.選択 = "分岐点貸付金" Then
            wsTbl = "DBDA010_分岐点貸付金"
            wsTbl2 = "DBDA010_分岐点貸付金"
        Else
            wsTbl = "DBDA010_貸付金"
        End If

        ws_Ginko = "外部貸付"

        Me.PageHeader.Controls("L_借入番号").Caption = "貸付番号"
        'Me.PageHeader.Controls("L_借入内容").Caption = "貸付内容"
        'Me.PageHeader.Controls("L_計画番号").Caption = "貸付計画番号"

    End Select
'
    'ラベル設定
    ws_Risoku = ""
    ws_Tyotan = ""
    
    '杉村倉庫仕様
'    wsTL(0) = "前払利息増"
'    wsTL(1) = "前払利息減"
'    wsTL(2) = "前払利息残高"
'    wsTL(3) = "未払利息増"
'    wsTL(4) = "未払利息減"
'    wsTL(5) = "未払利息残高"
'
'    wsTL(6) = "損益利息"
    wsTL(0) = "支払額"
    wsTL(1) = "前払利息(洗+)"
    wsTL(2) = "前払利息(計-)"
    wsTL(3) = "未払利息(洗-)"
    wsTL(4) = "未払利息(計+)"
    wsTL(5) = "支払利息"
    wsTL(6) = ""
'
    'ワーククリア
    For l = 0 To 4
        Cnt_Saki(l) = 0
        Cnt_Ato(l) = 0
        
        For j = 0 To wML '列数
            For k = 0 To wFD - 1 'Field数
                wd_Saki(l, j, k) = 0
                wd_Ato(l, j, k) = 0
            Next k
        Next j
    Next l
'
    'Top
    For j = 0 To wML
        ws01 = Right("00" + CStr(j), 2)
        '1
        Me.GroupFooter1.Controls("G11_" & ws01 & "1").Top = 330 '220
        Me.GroupFooter1.Controls("G11_" & ws01 & "2").Top = 550
        Me.GroupFooter1.Controls("G11_" & ws01 & "3").Top = 750
        Me.GroupFooter1.Controls("G11_" & ws01 & "4").Top = 970
        Me.GroupFooter1.Controls("G11_" & ws01 & "5").Top = 1190
        
        Me.GroupFooter1.Controls("G12_" & ws01 & "1").Top = 1630
        Me.GroupFooter1.Controls("G12_" & ws01 & "2").Top = 1850
        Me.GroupFooter1.Controls("G12_" & ws01 & "3").Top = 2070
        Me.GroupFooter1.Controls("G12_" & ws01 & "4").Top = 2290
        Me.GroupFooter1.Controls("G12_" & ws01 & "5").Top = 2510
        '2
        Me.GroupFooter2.Controls("G21_" & ws01 & "1").Top = 330
        Me.GroupFooter2.Controls("G21_" & ws01 & "2").Top = 550
        Me.GroupFooter2.Controls("G21_" & ws01 & "3").Top = 750
        Me.GroupFooter2.Controls("G21_" & ws01 & "4").Top = 970
        Me.GroupFooter2.Controls("G21_" & ws01 & "5").Top = 1190
        
        Me.GroupFooter2.Controls("G22_" & ws01 & "1").Top = 1630
        Me.GroupFooter2.Controls("G22_" & ws01 & "2").Top = 1850
        Me.GroupFooter2.Controls("G22_" & ws01 & "3").Top = 2070
        Me.GroupFooter2.Controls("G22_" & ws01 & "4").Top = 2290
        Me.GroupFooter2.Controls("G22_" & ws01 & "5").Top = 2510
        '3
        Me.GroupFooter3.Controls("G31_" & ws01 & "1").Top = 330
        Me.GroupFooter3.Controls("G31_" & ws01 & "2").Top = 550
        Me.GroupFooter3.Controls("G31_" & ws01 & "3").Top = 750
        Me.GroupFooter3.Controls("G31_" & ws01 & "4").Top = 970
        Me.GroupFooter3.Controls("G31_" & ws01 & "5").Top = 1190
        
        Me.GroupFooter3.Controls("G32_" & ws01 & "1").Top = 1630
        Me.GroupFooter3.Controls("G32_" & ws01 & "2").Top = 1850
        Me.GroupFooter3.Controls("G32_" & ws01 & "3").Top = 2070
        Me.GroupFooter3.Controls("G32_" & ws01 & "4").Top = 2290
        Me.GroupFooter3.Controls("G32_" & ws01 & "5").Top = 2510
        '4
        Me.GroupFooter4.Controls("G41_" & ws01 & "1").Top = 330
        Me.GroupFooter4.Controls("G41_" & ws01 & "2").Top = 550
        Me.GroupFooter4.Controls("G41_" & ws01 & "3").Top = 750
        Me.GroupFooter4.Controls("G41_" & ws01 & "4").Top = 970
        Me.GroupFooter4.Controls("G41_" & ws01 & "5").Top = 1190
        
        Me.GroupFooter4.Controls("G42_" & ws01 & "1").Top = 1630
        Me.GroupFooter4.Controls("G42_" & ws01 & "2").Top = 1850
        Me.GroupFooter4.Controls("G42_" & ws01 & "3").Top = 2070
        Me.GroupFooter4.Controls("G42_" & ws01 & "4").Top = 2290
        Me.GroupFooter4.Controls("G42_" & ws01 & "5").Top = 2510
        '
        Me.ReportFooter.Controls("G91_" & ws01 & "1").Top = 313
        Me.ReportFooter.Controls("G91_" & ws01 & "2").Top = 526
        Me.ReportFooter.Controls("G91_" & ws01 & "3").Top = 739
        Me.ReportFooter.Controls("G91_" & ws01 & "4").Top = 951
        Me.ReportFooter.Controls("G91_" & ws01 & "5").Top = 1134
        
        Me.ReportFooter.Controls("G92_" & ws01 & "1").Top = 1630
        Me.ReportFooter.Controls("G92_" & ws01 & "2").Top = 1843
        Me.ReportFooter.Controls("G92_" & ws01 & "3").Top = 2055
        Me.ReportFooter.Controls("G92_" & ws01 & "4").Top = 2268
        Me.ReportFooter.Controls("G92_" & ws01 & "5").Top = 2480
    Next j
'
    '----------------------------------------------------------------
    '                           ** 見出し **
    '----------------------------------------------------------------
    出力日 = Now
    企業名 = GCoName
    
    If G金利SM = True Then
        L_帳票名.Caption = " " & GRpt.選択 & GRpt.帳票名 & " - 金利SM " & GRpt.集計 & " " & GRpt.推移 & "- "
    Else
        L_帳票名.Caption = " " & GRpt.選択 & GRpt.帳票名 & " -" & GRpt.集計 & " " & GRpt.推移 & "- "
    End If
    
    'H00_最終実績年月 = ""
    'If Gコントロール.最終実績年月 > CDate("2001/01/01") Then
    '    H00_最終実績年月 = Format(Gコントロール.最終実績年月, Gfmt年月)
    'End If

    'H00_借入計画番号 = GRpt.借入
    H00_金融リストラ番号 = GRpt.金融
    
    If GRpt.千円単位 = 1 Then
        w分母 = 1000
        L_単位.Caption = "（千円単位）"
    Else
        w分母 = 1
        L_単位 = "（円単位）"
    End If
'
    'Top

    '----------------------------------------------------------------
    '          ** グループ ・ Where ・ OrderＢｙ **
    '----------------------------------------------------------------
    If GRpt.S_種別 = "分類1" Then
        GroupHeader1.DataField = "GrpFld_KShubetu"
    ElseIf GRpt.S_種別 = "分類2" Then
        GroupHeader2.DataField = "GrpFld_KShubetu"
    ElseIf GRpt.S_種別 = "分類3" Then
        GroupHeader3.DataField = "GrpFld_KShubetu"
    ElseIf GRpt.S_種別 = "分類4" Then
        GroupHeader4.DataField = "GrpFld_KShubetu"
    End If
    
    If GRpt.S_部門 = "分類1" Then
        GroupHeader1.DataField = "GrpFld_Bumon"
    ElseIf GRpt.S_部門 = "分類2" Then
        GroupHeader2.DataField = "GrpFld_Bumon"
    ElseIf GRpt.S_部門 = "分類3" Then
        GroupHeader3.DataField = "GrpFld_Bumon"
    ElseIf GRpt.S_部門 = "分類4" Then
        GroupHeader4.DataField = "GrpFld_Bumon"
    End If
    
    If GRpt.S_金融 = "分類1" Then
        GroupHeader1.DataField = "GrpFld_Kinyu"
    ElseIf GRpt.S_金融 = "分類2" Then
        GroupHeader2.DataField = "GrpFld_Kinyu"
    ElseIf GRpt.S_金融 = "分類3" Then
        GroupHeader3.DataField = "GrpFld_Kinyu"
    ElseIf GRpt.S_金融 = "分類4" Then
        GroupHeader4.DataField = "GrpFld_Kinyu"
    End If
    
    If GRpt.S_銀行 = "分類1" Then
        GroupHeader1.DataField = "GrpFld_Ginko"
    ElseIf GRpt.S_銀行 = "分類2" Then
        GroupHeader2.DataField = "GrpFld_Ginko"
    ElseIf GRpt.S_銀行 = "分類3" Then
        GroupHeader3.DataField = "GrpFld_Ginko"
    ElseIf GRpt.S_銀行 = "分類4" Then
        GroupHeader4.DataField = "GrpFld_Ginko"
    End If
    
    'Detail出力後の内部集計がおかしいので出力させている
    GroupHeader5.DataField = "I_借入番号"
    
    If GStr = "金利GR" Then
        GRet = 金利GR_CHECK
        If GRet > 1 Then
            If GRpt.S_金利 = "分類1" Then
                GroupHeader1.DataField = "GrpFld_KGroup"
            ElseIf GRpt.S_金利 = "分類2" Then
                GroupHeader2.DataField = "GrpFld_KGroup"
            ElseIf GRpt.S_金利 = "分類3" Then
                GroupHeader3.DataField = "GrpFld_KGroup"
            ElseIf GRpt.S_金利 = "分類4" Then
                GroupHeader4.DataField = "GrpFld_KGroup"
            End If
        End If
    End If

    '帳票指示
    wsS = ""
    If GRpt.借入金管理区分 = P8.FCDbl(XMXA020_区分("借入金管理区分", "決算用")) Then
        wsS = wsS & "帳票指示:決算用 "
    Else
        wsS = wsS & "帳票指示:管理用 "
    End If
    
    '計名セット、Shapeカラー
    If GroupHeader1.DataField = "GrpFld_KShubetu" Then
        wsS = wsS & "分類1:借入金種別 "
        Me.GroupFooter1.Controls("G10_計名").DataField = "I_借入金種別名"
        Me.GroupFooter1.Controls("Shape_11").BackColor = &HFFFFC0
        Me.GroupFooter1.Controls("Shape_12").BackColor = &HFFFFC0
    ElseIf GroupHeader1.DataField = "GrpFld_Bumon" Then
        wsS = wsS & "分類1:部門 "
        Me.GroupFooter1.Controls("G10_計名").DataField = "I_部門名"
        Me.GroupFooter1.Controls("Shape_11").BackColor = &HC0FFC0
        Me.GroupFooter1.Controls("Shape_12").BackColor = &HC0FFC0
    ElseIf GroupHeader1.DataField = "GrpFld_Kinyu" Then
        wsS = wsS & "分類1:金融機関 "
        Me.GroupFooter1.Controls("G10_計名").DataField = "I_金融機関名"
        Me.GroupFooter1.Controls("Shape_11").BackColor = &HE0E0E0
        Me.GroupFooter1.Controls("Shape_12").BackColor = &HE0E0E0
    ElseIf GroupHeader1.DataField = "GrpFld_Ginko" Then
        wsS = wsS & "分類1:銀行 "
        Me.GroupFooter1.Controls("G10_計名").DataField = "I_銀行名"
        Me.GroupFooter1.Controls("Shape_11").BackColor = &HC0FFFF
        Me.GroupFooter1.Controls("Shape_12").BackColor = &HC0FFFF
    ElseIf GroupHeader1.DataField = "GrpFld_KGroup" Then
        wsS = wsS & "分類1:金利G "
        Me.GroupFooter1.Controls("G10_計名").DataField = "I_金利グループ名"
        Me.GroupFooter1.Controls("Shape_11").BackColor = C_LGreen
        Me.GroupFooter1.Controls("Shape_12").BackColor = C_LGreen
    End If
    
    If GroupHeader2.DataField = "GrpFld_KShubetu" Then
        wsS = wsS & "分類2:借入金種別 "
        Me.GroupFooter2.Controls("G20_計名").DataField = "I_借入金種別名"
        Me.GroupFooter2.Controls("Shape_21").BackColor = &HFFFFC0
        Me.GroupFooter2.Controls("Shape_22").BackColor = &HFFFFC0
    ElseIf GroupHeader2.DataField = "GrpFld_Bumon" Then
        wsS = wsS & "分類2:部門 "
        Me.GroupFooter2.Controls("G20_計名").DataField = "I_部門名"
        Me.GroupFooter2.Controls("Shape_21").BackColor = &HC0FFC0
        Me.GroupFooter2.Controls("Shape_22").BackColor = &HC0FFC0
    ElseIf GroupHeader2.DataField = "GrpFld_Kinyu" Then
        wsS = wsS & "分類2:金融機関 "
        Me.GroupFooter2.Controls("G20_計名").DataField = "I_金融機関名"
        Me.GroupFooter2.Controls("Shape_21").BackColor = &HE0E0E0
        Me.GroupFooter2.Controls("Shape_22").BackColor = &HE0E0E0
    ElseIf GroupHeader2.DataField = "GrpFld_Ginko" Then
        wsS = wsS & "分類2:銀行 "
        Me.GroupFooter2.Controls("G20_計名").DataField = "I_銀行名"
        Me.GroupFooter2.Controls("Shape_21").BackColor = &HC0FFFF
        Me.GroupFooter2.Controls("Shape_22").BackColor = &HC0FFFF
    ElseIf GroupHeader2.DataField = "GrpFld_KGroup" Then
        wsS = wsS & "分類2:金利G "
        Me.GroupFooter2.Controls("G20_計名").DataField = "I_金利グループ名"
        Me.GroupFooter2.Controls("Shape_21").BackColor = C_LGreen
        Me.GroupFooter2.Controls("Shape_22").BackColor = C_LGreen
    End If
    
    If GroupHeader3.DataField = "GrpFld_KShubetu" Then
        wsS = wsS & "分類3:借入金種別 "
        Me.GroupFooter3.Controls("G30_計名").DataField = "I_借入金種別名"
        Me.GroupFooter3.Controls("Shape_31").BackColor = &HFFFFC0
        Me.GroupFooter3.Controls("Shape_32").BackColor = &HFFFFC0
    ElseIf GroupHeader3.DataField = "GrpFld_Bumon" Then
        wsS = wsS & "分類3:部門 "
        Me.GroupFooter3.Controls("G30_計名").DataField = "I_部門名"
        Me.GroupFooter3.Controls("Shape_31").BackColor = &HC0FFC0
        Me.GroupFooter3.Controls("Shape_32").BackColor = &HC0FFC0
    ElseIf GroupHeader3.DataField = "GrpFld_Kinyu" Then
        wsS = wsS & "分類3:金融機関 "
        Me.GroupFooter3.Controls("G30_計名").DataField = "I_金融機関名"
        Me.GroupFooter3.Controls("Shape_31").BackColor = &HE0E0E0
        Me.GroupFooter3.Controls("Shape_32").BackColor = &HE0E0E0
    ElseIf GroupHeader3.DataField = "GrpFld_Ginko" Then
        wsS = wsS & "分類3:銀行 "
        Me.GroupFooter3.Controls("G30_計名").DataField = "I_銀行名"
        Me.GroupFooter3.Controls("Shape_31").BackColor = &HC0FFFF
        Me.GroupFooter3.Controls("Shape_32").BackColor = &HC0FFFF
    ElseIf GroupHeader3.DataField = "GrpFld_KGroup" Then
        wsS = wsS & "分類3:金利G "
        Me.GroupFooter3.Controls("G30_計名").DataField = "I_金利グループ名"
        Me.GroupFooter3.Controls("Shape_31").BackColor = C_LGreen
        Me.GroupFooter3.Controls("Shape_32").BackColor = C_LGreen
    End If
    
    If GroupHeader4.DataField = "GrpFld_KShubetu" Then
        wsS = wsS & "分類4:借入金種別 "
        Me.GroupFooter4.Controls("G40_計名").DataField = "I_借入金種別名"
        Me.GroupFooter4.Controls("Shape_41").BackColor = &HFFFFC0
        Me.GroupFooter4.Controls("Shape_42").BackColor = &HFFFFC0
    ElseIf GroupHeader4.DataField = "GrpFld_Bumon" Then
        wsS = wsS & "分類4:部門 "
        Me.GroupFooter4.Controls("G40_計名").DataField = "I_部門名"
        Me.GroupFooter4.Controls("Shape_41").BackColor = &HC0FFC0
        Me.GroupFooter4.Controls("Shape_42").BackColor = &HC0FFC0
    ElseIf GroupHeader4.DataField = "GrpFld_Kinyu" Then
        wsS = wsS & "分類4:金融機関 "
        Me.GroupFooter4.Controls("G40_計名").DataField = "I_金融機関名"
        Me.GroupFooter4.Controls("Shape_41").BackColor = &HE0E0E0
        Me.GroupFooter4.Controls("Shape_42").BackColor = &HE0E0E0
    ElseIf GroupHeader4.DataField = "GrpFld_Ginko" Then
        wsS = wsS & "分類4:銀行 "
        Me.GroupFooter4.Controls("G40_計名").DataField = "I_銀行名"
        Me.GroupFooter4.Controls("Shape_41").BackColor = &HC0FFFF
        Me.GroupFooter4.Controls("Shape_42").BackColor = &HC0FFFF
    ElseIf GroupHeader4.DataField = "GrpFld_KGroup" Then
        wsS = wsS & "分類4:金利G "
        Me.GroupFooter4.Controls("G40_計名").DataField = "I_金利グループ名"
        Me.GroupFooter4.Controls("Shape_41").BackColor = C_LGreen
        Me.GroupFooter4.Controls("Shape_42").BackColor = C_LGreen
    End If
'
    '杉村倉庫仕様
    Me.GroupFooter1.Controls("L_G112").Caption = wsTL(0)
    Me.GroupFooter1.Controls("L_G113").Caption = wsTL(1)
    Me.GroupFooter1.Controls("L_G114").Caption = wsTL(2)
    Me.GroupFooter1.Controls("L_G115").Caption = wsTL(5)
    Me.GroupFooter1.Controls("L_G122").Caption = wsTL(0)
    Me.GroupFooter1.Controls("L_G123").Caption = wsTL(3)
    Me.GroupFooter1.Controls("L_G124").Caption = wsTL(4)
    Me.GroupFooter1.Controls("L_G125").Caption = wsTL(5)

    Me.GroupFooter2.Controls("L_G212").Caption = wsTL(0)
    Me.GroupFooter2.Controls("L_G213").Caption = wsTL(1)
    Me.GroupFooter2.Controls("L_G214").Caption = wsTL(2)
    Me.GroupFooter2.Controls("L_G215").Caption = wsTL(5)
    Me.GroupFooter2.Controls("L_G222").Caption = wsTL(0)
    Me.GroupFooter2.Controls("L_G223").Caption = wsTL(3)
    Me.GroupFooter2.Controls("L_G224").Caption = wsTL(4)
    Me.GroupFooter2.Controls("L_G225").Caption = wsTL(5)

    Me.GroupFooter3.Controls("L_G312").Caption = wsTL(0)
    Me.GroupFooter3.Controls("L_G313").Caption = wsTL(1)
    Me.GroupFooter3.Controls("L_G314").Caption = wsTL(2)
    Me.GroupFooter3.Controls("L_G315").Caption = wsTL(5)
    Me.GroupFooter3.Controls("L_G322").Caption = wsTL(0)
    Me.GroupFooter3.Controls("L_G323").Caption = wsTL(3)
    Me.GroupFooter3.Controls("L_G324").Caption = wsTL(4)
    Me.GroupFooter3.Controls("L_G325").Caption = wsTL(5)

    Me.GroupFooter4.Controls("L_G412").Caption = wsTL(0)
    Me.GroupFooter4.Controls("L_G413").Caption = wsTL(1)
    Me.GroupFooter4.Controls("L_G414").Caption = wsTL(2)
    Me.GroupFooter4.Controls("L_G415").Caption = wsTL(5)
    Me.GroupFooter4.Controls("L_G422").Caption = wsTL(0)
    Me.GroupFooter4.Controls("L_G423").Caption = wsTL(3)
    Me.GroupFooter4.Controls("L_G424").Caption = wsTL(4)
    Me.GroupFooter4.Controls("L_G425").Caption = wsTL(5)

    Me.ReportFooter.Controls("L_G912").Caption = wsTL(0)
    Me.ReportFooter.Controls("L_G913").Caption = wsTL(1)
    Me.ReportFooter.Controls("L_G914").Caption = wsTL(2)
    Me.ReportFooter.Controls("L_G915").Caption = wsTL(5)
    Me.ReportFooter.Controls("L_G922").Caption = wsTL(0)
    Me.ReportFooter.Controls("L_G923").Caption = wsTL(3)
    Me.ReportFooter.Controls("L_G924").Caption = wsTL(4)
    Me.ReportFooter.Controls("L_G925").Caption = wsTL(5)
'
    GroupFooter1.Visible = True
    GroupFooter2.Visible = True
    GroupFooter3.Visible = True
    GroupFooter4.Visible = True
    If GroupHeader1.DataField = "" Then
        GroupFooter1.Visible = False
    End If
    If GroupHeader2.DataField = "" Then
        GroupFooter2.Visible = False
    End If
    If GroupHeader3.DataField = "" Then
        GroupFooter3.Visible = False
    End If
    If GroupHeader4.DataField = "" Then
        GroupFooter4.Visible = False
    End If
'
    '総合計の表示/非表示
    ReportFooter.Visible = True
    'If GRpt.指定 <> "" Then
    '    ReportFooter.Visible = False
    'End If
    If GRpt.C_種別 <> "" Then
        If GRpt.S_種別 = "分類1" Then
            ReportFooter.Visible = False
        End If
        wsS = wsS & "借入金種別名:" & GRpt.C_種別 & " "
    End If
    
    If GRpt.C_部門 <> "" Then
        If GRpt.S_部門 = "分類1" Then
            ReportFooter.Visible = False
        End If
        wsS = wsS & "部門名:" & GRpt.C_部門 & " "
    End If

    
    If GRpt.C_金融 <> "" Then
        If GRpt.S_金融 = "分類1" Then
            ReportFooter.Visible = False
        End If
        wsS = wsS & "金融機関名:" & GRpt.C_金融 & " "
    End If
    
    If GRpt.C_銀行 <> "" Then
        If GRpt.S_銀行 = "分類1" Then
            ReportFooter.Visible = False
        End If
        wsS = wsS & "銀行名:" & GRpt.C_銀行 & " "
    End If
    
    '帳票指示
    Me.PageHeader.Controls("L_帳票指示").Caption = wsS
    
    '改ページ
    Me.GroupFooter1.NewPage = GRpt.NewPage1
    Me.GroupFooter2.NewPage = GRpt.NewPage2
    Me.GroupFooter3.NewPage = GRpt.NewPage3
    Me.GroupFooter4.NewPage = GRpt.NewPage4
    
    '印刷設定
    If GRpt.詳細表示 = 1 Then
        Me.Detail.Height = 1340
    Else
        Me.Detail.Height = 0
    End If
    
    'wWhere = ""
    'wWhere = wWhere & " Where (1=1) "
    '銀行指定
    'If GRpt.指定 <> "" Then
    '    wWhere = wWhere & " And G.銀行名='" & GRpt.指定 & "'"
    'End If
    
    'Order
    If GStr <> "金利GR" Then
        wWhere = wWhere & " ORDER BY K.借入金種別区分,K.銀行番号,Z.借入番号"
        'wWhere = wWhere & " ORDER By K.銀行番号,借入金種別区分,Z.借入番号"
    Else
        '金利SM
        wWhere = wWhere & " ORDER BY IIF(K.金利グループ区分<>'',K.金利グループ区分,'99999'),K.銀行番号,Z.借入番号"
    End If
    
    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    'Call MRB010_標準入力借入残高表(wsTbl, wsTbl2)       '07/02/18 V180
    Call MBD020_借入金ワークテーブル作成(wsTbl) 'データ絞り込み
    Call MRB010_標準入力借入残高表("DCIA010_借入金ワーク")  '16/03/26 利子補給に伴う変更
    'Call MRB010_標準入力借入残高表固定日数("DCIA010_借入金ワーク")
    Call MRB010_手入力借入残高表("DCIA010_借入金ワーク")         '07/02/09 V180
'
    'w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "平成")
    '2019/01/15 日付入力区分 仕様変更
    If G基本情報.日付入力区分 = "0" Then
    '和暦
        If Len(GRpt.テキスト_01) <= 2 Then
            w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "平成")
        Else
        w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "西暦")
        End If
    Else
    '西暦
        w開始年月日 = C年月日.年度開始年月日(GRpt.テキスト_01, "西暦")
    End If
    
    w推移表タイトル = MUA010_推移表ファイル作成("", "", w開始年月日, GRpt.推移, wML)
    
    Call RDB010_コントロールセット
    
    '----------------------------------------------------------------
    '                          ** 合計 計算 **
    '----------------------------------------------------------------
    '----------------------------------------------------------------
    '                       ** レコード　ソース **
    '----------------------------------------------------------------
    '** レコード　ソース **
    wstr = "Select "
    wstr = wstr & "K.借入番号 As I_借入番号,"
    
    'セクションGR
    wstr = wstr & "K.銀行番号 As GrpFld_Ginko,"
    wstr = wstr & "G.金融機関番号 As GrpFld_Kinyu,"
    wstr = wstr & "B.部門番号 As GrpFld_Bumon,"
    wstr = wstr & " K.借入金種別区分 As GrpFld_KShubetu,"
    If GStr = "金利GR" Then
        wstr = wstr & "IIF(K.金利グループ区分<>'',K.金利グループ区分,'99999') As GrpFld_KGroup,"
    End If
    
    wstr = wstr & "G.銀行名 As I_銀行名,"
    wstr = wstr & "G.金融機関名 As I_金融機関名,"
    wstr = wstr & "B.部門名 As I_部門名,"
    wstr = wstr & "S.借入金種別名 As I_借入金種別名,"
    If GStr <> "金利GR" Then
        wstr = wstr & "'' As I_金利グループ名,"
    Else
        wstr = wstr & "IIF(KG.金利グループ名<>'',KG.金利グループ名,'グループ無') As I_金利グループ名,"
    End If
    
    wstr = wstr & "K.sm区分 As I_SM区分,"
    wstr = wstr & "K.金融リストラ番号 As I_金融リストラ番号,"
    wstr = wstr & "Format(K.金融解約実行日,'" & Gfmt年月日 & "') As I_金融解約年月日,"
    
    
    'Order Grp
    wstr = wstr & " IIF(K.長短区分 = " & P8.FCDbl(XMXA020_区分("長短区分", "短期借入金")) & ",'短期借入金','長期借入金') AS I_長短区分,"
    'wstr = wstr & " IIF(K.有担保フラグ = " & P8.FCDbl(XMXA020_区分("有担フラグ", "無担保")) & ",'無担保','有担保') As I_担保,"
    'wstr = wstr & " IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As I_金利種別,"
    
    'wstr = wstr & "G.銀行名 As I_銀行名,"
    
    'wstr = wstr & "Format(K.解約実行日,'" & Gfmt年月日 & "') As I_解約年月日,"
    'wstr = wstr & "K.借入計画番号 As I_借入計画番号,"
    wstr = wstr & " IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As I_利息区分名,"
    
    wstr = wstr & "Z.元金_01+解約_01+Z.残高_01 As I_期首残高,"
    
    For j = 1 To wML '列数
        w番号 = Right("00" & CStr(j), 2)

        wstr = wstr & "Z.残高_" & w番号 & " As I_" & w番号 & "1,"
'        wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息増_" & w番号 & ",Z.未払利息増_" & w番号 & ") As I_" & w番号 & "2,"
'        wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & ",Z.未払利息減_" & w番号 & ") As I_" & w番号 & "3,"
'        wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息_" & w番号 & ",Z.未払利息_" & w番号 & ") As I_" & w番号 & "4,"
'        wstr = wstr & "Z2.損益利息額_" & w番号 & " As I_" & w番号 & "5,"
        '杉村倉庫仕様
        wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息増_" & w番号 & ",Z.未払利息減_" & w番号 & ") As I_" & w番号 & "2,"
        wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & "-Z.前払利息増_" & w番号 & "+Z.前払利息_" & w番号 & ",-Z.未払利息増_" & w番号 & "+Z.未払利息減_" & w番号 & "+Z.未払利息_" & w番号 & ") As I_" & w番号 & "3,"
        wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息_" & w番号 & ",Z.未払利息_" & w番号 & ") As I_" & w番号 & "4,"
        wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & ",Z.未払利息増_" & w番号 & ") As I_" & w番号 & "5,"
    Next
    
    wstr = wstr & "Z.残高合計 As I_001,"
'    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息増合計,Z.未払利息増合計) As I_002,"
'    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減合計,Z.未払利息減合計) As I_003,"
'    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息合計,Z.未払利息合計) As I_004,"
'    wstr = wstr & "Z2.損益利息額合計 As I_005"
    '杉村倉庫仕様
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息増合計,Z.未払利息減合計) As I_002,"
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減合計-Z.前払利息増合計+Z.前払利息合計,-Z.未払利息増合計+Z.未払利息減合計+Z.未払利息合計) As I_003,"
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息合計,Z.未払利息合計) As I_004,"
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減合計,Z.未払利息増合計) As I_005"
    
    wstr = wstr & " FROM (((((DCDA010_借入残高推移表結果 As Z"
    wstr = wstr & " INNER JOIN DCDA010_借入残高推移表結果２ As Z2"
    wstr = wstr & " ON Z.借入番号=Z2.借入番号)"
    wstr = wstr & " INNER JOIN DCIA010_借入金ワーク As K"
    wstr = wstr & " ON Z.借入番号=K.借入番号)"
    wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr & " ON K.銀行番号=G.銀行番号)"
    wstr = wstr & " LEFT JOIN DAAA200_部門マスタ As B"
    wstr = wstr & " ON K.プロジェクト番号 = B.部門番号)"
    wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
    wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分)"
'    wstr = wstr & " LEFT JOIN DAAA115_金利シミュレーショングループ As KG"
'    wstr = wstr & " ON K.金利グループ区分 = KG.金利グループ区分"
    
    If GStr = "金利GR" Then
        wstr = wstr & " LEFT JOIN DAAA115_金利シミュレーショングループ As KG"
        wstr = wstr & " ON K.金利グループ区分 = KG.金利グループ区分"
    End If
    
    'Where 条件
    'All 0 は表示しない
    wstr = wstr & " Where (Z.前払利息合計<>0"
    wstr = wstr & " Or Z.前払利息増合計<>0"
    wstr = wstr & " Or Z.前払利息減合計<>0"
    wstr = wstr & " Or Z.未払利息合計<>0"
    wstr = wstr & " Or Z.未払利息増合計<>0"
    wstr = wstr & " Or Z.未払利息減合計<>0"
    wstr = wstr & " Or Z2.損益利息額合計<>0"
    
    wstr = wstr & " Or ("
    For j = 1 To wML - 1 '列数
        w番号 = Right("00" & CStr(j), 2)

        wstr = wstr & "Z.残高_" & w番号 & "<>0 Or "
    Next
    w番号 = Right("00" & CStr(wML), 2)
    wstr = wstr & "Z.残高_" & w番号 & "<>0"
    wstr = wstr & "))"
    
    '銀行指定
    'If GRpt.指定 <> "" Then
    '    wstr = wstr & " And G.銀行名='" & GRpt.指定 & "'"
    'End If
    
    'Order
'    If GStr <> "金利GR" Then
'        wstr = wstr & " ORDER BY K.借入金種別区分,K.銀行番号,1,2,K.借入番号"
'        'wstr = wstr & " ORDER BY K.銀行番号,K.借入金種別区分,1,2,K.借入番号"
'    Else
'        '金利SM
'        wstr = wstr & " ORDER BY IIF(K.金利グループ区分<>'',K.金利グループ区分,'99999'),K.銀行番号,Z.借入番号"
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
    
    Me.DataControl1.Source = wstr
'
    Exit Sub
'
'----------< ERROR ROUTINE >--------------------------------------------------------
ActiveReport_ReportStart_ERR:
    pERR_MES = pPROGRAM_ID + "/ ActiveReport_ReportStart() でエラー" + vbCrLf + vbCrLf + _
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
' ActiveReport_ReportEnd
'------------------------------------------------
Private Sub ActiveReport_ReportEnd()
'
    'FBA010_帳票範囲指定.メッセージ = ""
    'FBA010_帳票範囲指定.メッセージ.Refresh
'
    ' =========================================
    '           　 CsvFile 作成
    ' =========================================
    If GRpt.CSV = 1 Then
        Call MX040_CsvOut_KARISUII(w推移表タイトル)
    End If
'
    ' =========================================
    '           　 ボタン制御
    ' =========================================
    'FBA010_帳票範囲指定.実行.Enabled = True
    'FBA010_帳票範囲指定.閉じる.Enabled = True
'
    'FBA010_帳票範囲指定.拡張.SetFocus
'
    ' =========================================
    '  借換たろう！お試し版帳票出力回数チェック
    ' =========================================
    If GSys.Sys = "借入金 お試し版" Then
        Call MAA001_KARIKAETAROU_CNT
    End If
'
End Sub

'------------------------------------------------
' ActiveReport_NoData
'------------------------------------------------
Private Sub ActiveReport_NoData()
'
    'FBA010_帳票範囲指定.メッセージ = "出力すべきデータはありません"
    'FBA010_帳票範囲指定.メッセージ.Refresh
    GSstrt帳票Msg = "出力すべきデータはありません"
'
    Me.Cancel
    DoEvents
'
    ' =========================================
    '           　 ボタン制御
    ' =========================================
    'FBA010_帳票範囲指定.実行.Enabled = True
    'FBA010_帳票範囲指定.閉じる.Enabled = True
'
    'FBA010_帳票範囲指定.拡張.SetFocus
'
    Unload Me
'
End Sub

'------------------------------------------------
' ActiveReport_Error
'------------------------------------------------
Private Sub ActiveReport_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
'
    'FBA010_帳票範囲指定.メッセージ = "出力できませんでした"
    'FBA010_帳票範囲指定.メッセージ.Refresh
    GSstrt帳票Msg = "出力できませんでした"
'
    Me.Cancel
    DoEvents

    ' =========================================
    '           　 ボタン制御
    ' =========================================
    'FBA010_帳票範囲指定.実行.Enabled = True
    'FBA010_帳票範囲指定.閉じる.Enabled = True
'
    'FBA010_帳票範囲指定.拡張.SetFocus
'
    Unload Me
'
End Sub

'------------------------------------------------
' RDB010_コントロールセット
'------------------------------------------------
Private Sub RDB010_コントロールセット()
    Dim j As Integer
'
    For j = 1 To wML '列数
        w番号 = Right("00" + CStr(j), 2)
        Me.PageHeader.Controls("Lbl_" + w番号 + "番目年月") = w推移表タイトル.X番目年月(j)
    Next
'
End Sub

'------------------------------------------------
' Detail_BeforePrint
'------------------------------------------------
Private Sub Detail_BeforePrint()
'
    Dim j As Integer, k As Integer
    Dim wd01 As Double, wd02 As Double
    Dim wstr As String
    Dim wism As Integer
'
    wism = P8.FCDbl(Me.Detail.Controls("I_SM区分"))
    
    Me.Detail.Controls("I_SM区分") = ""
    If wism = 1 And P8.FCStr(Me.Detail.Controls("I_金融リストラ番号")) <> "" Then
        Me.Detail.Controls("I_SM区分") = "借入ｼﾐｭﾚｰｼｮﾝ"
    ElseIf wism = 0 And P8.FCStr(Me.Detail.Controls("I_金融解約年月日")) <> "" Then
        Me.Detail.Controls("I_SM区分") = "解約ｼﾐｭﾚｰｼｮﾝ"
    End If
'
    ws_Risoku = P8.FCStr(Me.Detail.Controls("I_利息区分名"))
    ws_Tyotan = P8.FCStr(Me.Detail.Controls("I_長短区分"))
    
'    If ws_Risoku = "利息先払" Then
'        Me.Detail.Controls("L_D02").Caption = wsTL(0)
'        Me.Detail.Controls("L_D03").Caption = wsTL(1)
'        Me.Detail.Controls("L_D04").Caption = wsTL(2)
'    Else
'        Me.Detail.Controls("L_D02").Caption = wsTL(3)
'        Me.Detail.Controls("L_D03").Caption = wsTL(4)
'        Me.Detail.Controls("L_D04").Caption = wsTL(5)
'    End If
'
'    Me.Detail.Controls("L_D05").Caption = wsTL(6)
    '杉村倉庫仕様
    Me.Detail.Controls("L_D02").Caption = wsTL(0)
    If ws_Risoku = "利息先払" Then
        Me.Detail.Controls("L_D03").Caption = wsTL(1)
        Me.Detail.Controls("L_D04").Caption = wsTL(2)
    Else
        Me.Detail.Controls("L_D03").Caption = wsTL(3)
        Me.Detail.Controls("L_D04").Caption = wsTL(4)
    End If
    Me.Detail.Controls("L_D05").Caption = wsTL(5)
'
    For j = 0 To wML '列数
        For k = 1 To wFD 'Field数
            wstr = Right("00" + CStr(j), 2) + CStr(k)
            Call MXA030_ReportColor(Me.Detail.Controls("I_" + wstr))
            wd01 = P8.FCDblRD(Me.Detail.Controls("I_" + wstr))
            Me.Detail.Controls("I_" + wstr) = Format(wd01 / w分母, "#,##0")
            
            '小計 総合計 集計
            If ws_Risoku = "利息先払" Then
                wd_Saki(0, j, k - 1) = wd_Saki(0, j, k - 1) + wd01
                wd_Saki(1, j, k - 1) = wd_Saki(1, j, k - 1) + wd01
                wd_Saki(2, j, k - 1) = wd_Saki(2, j, k - 1) + wd01
                wd_Saki(3, j, k - 1) = wd_Saki(3, j, k - 1) + wd01
                wd_Saki(4, j, k - 1) = wd_Saki(4, j, k - 1) + wd01
            ElseIf ws_Risoku = "利息後払" Then
                wd_Ato(0, j, k - 1) = wd_Ato(0, j, k - 1) + wd01
                wd_Ato(1, j, k - 1) = wd_Ato(1, j, k - 1) + wd01
                wd_Ato(2, j, k - 1) = wd_Ato(2, j, k - 1) + wd01
                wd_Ato(3, j, k - 1) = wd_Ato(3, j, k - 1) + wd01
                wd_Ato(4, j, k - 1) = wd_Ato(4, j, k - 1) + wd01
            End If
        Next k
    Next j
'
    '小計 総合計 集計
    If ws_Risoku = "利息先払" Then
        Cnt_Saki(0) = Cnt_Saki(0) + 1
        Cnt_Saki(1) = Cnt_Saki(1) + 1
        Cnt_Saki(2) = Cnt_Saki(2) + 1
        Cnt_Saki(3) = Cnt_Saki(3) + 1
        Cnt_Saki(4) = Cnt_Saki(4) + 1
    ElseIf ws_Risoku = "利息後払" Then
        Cnt_Ato(0) = Cnt_Ato(0) + 1
        Cnt_Ato(1) = Cnt_Ato(1) + 1
        Cnt_Ato(2) = Cnt_Ato(2) + 1
        Cnt_Ato(3) = Cnt_Ato(3) + 1
        Cnt_Ato(4) = Cnt_Ato(4) + 1
    End If
'
End Sub

'------------------------------------------------
' GroupFooter1_BeforePrint
'------------------------------------------------
Private Sub GroupFooter1_BeforePrint()
'
    Dim j As Integer, k As Integer
    Dim wd01 As Double
    Dim wstr As String
'
    wstr = Me.GroupFooter1.Controls("G10_計名")
    Me.GroupFooter1.Controls("G11_計名") = wstr & " 前払費用計"
    Me.GroupFooter1.Controls("G12_計名") = wstr & " 未払費用計"

    Me.GroupFooter1.Controls("G11_件数") = Format(Cnt_Saki(0), "#,##0")
    Me.GroupFooter1.Controls("G12_件数") = Format(Cnt_Ato(0), "#,##0")
'
    For j = 0 To wML '列数
        For k = 1 To wFD 'Field数
            wstr = Right("00" + CStr(j), 2) + CStr(k)
            
            Me.GroupFooter1.Controls("G11_" + wstr) = Format(wd_Saki(0, j, k - 1) / w分母, "#,##0")
            Call MXA030_ReportColor(Me.GroupFooter1.Controls("G11_" + wstr))
            
            Me.GroupFooter1.Controls("G12_" + wstr) = Format(wd_Ato(0, j, k - 1) / w分母, "#,##0")
            Call MXA030_ReportColor(Me.GroupFooter1.Controls("G12_" + wstr))
            
        Next k
    Next j
'
End Sub

'------------------------------------------------
' GroupFooter1_AfterPrint
'------------------------------------------------
Private Sub GroupFooter1_AfterPrint()
'
    Dim j As Integer, k As Integer
'
    '小計 クリア
    Cnt_Saki(0) = 0
    Cnt_Ato(0) = 0
    For j = 0 To wML '列数
        For k = 0 To wFD - 1 'Field数
            wd_Saki(0, j, k) = 0
            wd_Ato(0, j, k) = 0
        Next k
    Next j
'
End Sub

'------------------------------------------------
' GroupFooter2_BeforePrint
'------------------------------------------------
Private Sub GroupFooter2_BeforePrint()
'
    Dim j As Integer, k As Integer
    Dim wd01 As Double
    Dim wstr As String
'
    wstr = Me.GroupFooter2.Controls("G20_計名")
    Me.GroupFooter2.Controls("G21_計名") = wstr & " 前払費用計"
    Me.GroupFooter2.Controls("G22_計名") = wstr & " 未払費用計"

    Me.GroupFooter2.Controls("G21_件数") = Format(Cnt_Saki(1), "#,##0")
    Me.GroupFooter2.Controls("G22_件数") = Format(Cnt_Ato(1), "#,##0")
'
    For j = 0 To wML '列数
        For k = 1 To wFD 'Field数
            wstr = Right("00" + CStr(j), 2) + CStr(k)
            
            Me.GroupFooter2.Controls("G21_" + wstr) = Format(wd_Saki(1, j, k - 1) / w分母, "#,##0")
            Call MXA030_ReportColor(Me.GroupFooter2.Controls("G21_" + wstr))
            
            Me.GroupFooter2.Controls("G22_" + wstr) = Format(wd_Ato(1, j, k - 1) / w分母, "#,##0")
            Call MXA030_ReportColor(Me.GroupFooter2.Controls("G22_" + wstr))
            
        Next k
    Next j
'
End Sub

'------------------------------------------------
' GroupFooter2_AfterPrint
'------------------------------------------------
Private Sub GroupFooter2_AfterPrint()
'
    Dim j As Integer, k As Integer
'
    '小計 クリア
    Cnt_Saki(1) = 0
    Cnt_Ato(1) = 0
    For j = 0 To wML '列数
        For k = 0 To wFD - 1 'Field数
            wd_Saki(1, j, k) = 0
            wd_Ato(1, j, k) = 0
        Next k
    Next j
'
End Sub

'------------------------------------------------
' GroupFooter3_BeforePrint
'------------------------------------------------
Private Sub GroupFooter3_BeforePrint()
'
    Dim j As Integer, k As Integer
    Dim wd01 As Double
    Dim wstr As String
'
    wstr = Me.GroupFooter3.Controls("G30_計名")
    Me.GroupFooter3.Controls("G31_計名") = wstr & " 前払費用計"
    Me.GroupFooter3.Controls("G32_計名") = wstr & " 未払費用計"

    Me.GroupFooter3.Controls("G31_件数") = Format(Cnt_Saki(2), "#,##0")
    Me.GroupFooter3.Controls("G32_件数") = Format(Cnt_Ato(2), "#,##0")
'
    For j = 0 To wML '列数
        For k = 1 To wFD 'Field数
            wstr = Right("00" + CStr(j), 2) + CStr(k)
            
            Me.GroupFooter3.Controls("G31_" + wstr) = Format(wd_Saki(2, j, k - 1) / w分母, "#,##0")
            Call MXA030_ReportColor(Me.GroupFooter3.Controls("G31_" + wstr))
            
            Me.GroupFooter3.Controls("G32_" + wstr) = Format(wd_Ato(2, j, k - 1) / w分母, "#,##0")
            Call MXA030_ReportColor(Me.GroupFooter3.Controls("G32_" + wstr))
            
        Next k
    Next j
'
End Sub

'------------------------------------------------
' GroupFooter3_AfterPrint
'------------------------------------------------
Private Sub GroupFooter3_AfterPrint()
'
    Dim j As Integer, k As Integer
'
    '小計 クリア
    Cnt_Saki(2) = 0
    Cnt_Ato(2) = 0
    For j = 0 To wML '列数
        For k = 0 To wFD - 1 'Field数
            wd_Saki(2, j, k) = 0
            wd_Ato(2, j, k) = 0
        Next k
    Next j
'
End Sub

'------------------------------------------------
' GroupFooter4_BeforePrint
'------------------------------------------------
Private Sub GroupFooter4_BeforePrint()
'
    Dim j As Integer, k As Integer
    Dim wd01 As Double
    Dim wstr As String
'
    wstr = Me.GroupFooter4.Controls("G40_計名")
    Me.GroupFooter4.Controls("G41_計名") = wstr & " 前払費用計"
    Me.GroupFooter4.Controls("G42_計名") = wstr & " 未払費用計"

    Me.GroupFooter4.Controls("G41_件数") = Format(Cnt_Saki(3), "#,##0")
    Me.GroupFooter4.Controls("G42_件数") = Format(Cnt_Ato(3), "#,##0")
'
    For j = 0 To wML '列数
        For k = 1 To wFD 'Field数
            wstr = Right("00" + CStr(j), 2) + CStr(k)
            
            Me.GroupFooter4.Controls("G41_" + wstr) = Format(wd_Saki(3, j, k - 1) / w分母, "#,##0")
            Call MXA030_ReportColor(Me.GroupFooter4.Controls("G41_" + wstr))
            
            Me.GroupFooter4.Controls("G42_" + wstr) = Format(wd_Ato(3, j, k - 1) / w分母, "#,##0")
            Call MXA030_ReportColor(Me.GroupFooter4.Controls("G42_" + wstr))
            
        Next k
    Next j
'
End Sub

'------------------------------------------------
' GroupFooter4_AfterPrint
'------------------------------------------------
Private Sub GroupFooter4_AfterPrint()
'
    Dim j As Integer, k As Integer
'
    '小計 クリア
    Cnt_Saki(3) = 0
    Cnt_Ato(3) = 0
    For j = 0 To wML '列数
        For k = 0 To wFD - 1 'Field数
            wd_Saki(3, j, k) = 0
            wd_Ato(3, j, k) = 0
        Next k
    Next j
'
End Sub

'------------------------------------------------
' ReportFooter_BeforePrint
'------------------------------------------------
Private Sub ReportFooter_BeforePrint()
'
    Dim j As Integer, k As Integer
    Dim wd01 As Double
    Dim wstr As String
'
    Me.ReportFooter.Controls("G91_件数") = Format(Cnt_Saki(4), "#,##0")
    Me.ReportFooter.Controls("G92_件数") = Format(Cnt_Ato(4), "#,##0")
'
    For j = 0 To wML '列数
        For k = 1 To 5 'Field数
            wstr = Right("00" + CStr(j), 2) + CStr(k)
            
            Me.ReportFooter.Controls("G91_" + wstr) = Format(wd_Saki(4, j, k - 1) / w分母, "#,##0")
            Call MXA030_ReportColor(Me.ReportFooter.Controls("G91_" + wstr))
            
            Me.ReportFooter.Controls("G92_" + wstr) = Format(wd_Ato(4, j, k - 1) / w分母, "#,##0")
            Call MXA030_ReportColor(Me.ReportFooter.Controls("G92_" + wstr))
        Next k
    Next j
'
End Sub

'------------------------------------------------
' 金利GR_CHECK
'------------------------------------------------
Private Function 金利GR_CHECK() As Integer
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
'
    On Error GoTo 金利GR_CHECK_ERR
'
    wstr = "SELECT K.金利グループ区分"
    wstr = wstr & " FROM (DCDA010_借入残高推移表結果 AS Z"
    wstr = wstr & " INNER JOIN DCIA010_借入金ワーク AS K ON Z.借入番号 = K.借入番号)"
    wstr = wstr & " LEFT JOIN DAAA115_金利シミュレーショングループ AS KG"
    wstr = wstr & " ON K.金利グループ区分 = KG.金利グループ区分"
    wstr = wstr & " GROUP BY K.金利グループ区分"
    wstr = wstr & " Having K.金利グループ区分<>''"
    wstr = wstr & " ORDER BY K.金利グループ区分"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        金利GR_CHECK = wRs.RecordCount
    wRs.Close
    Set wRs = Nothing
'
    Exit Function
'
'----------< ERROR ROUTINE >--------------------------------------------------------
金利GR_CHECK_ERR:
    pERR_MES = pPROGRAM_ID + "/ 金利GR_CHECK() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Function


