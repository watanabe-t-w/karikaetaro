VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RDF011_利息前払未払残高表 
   Caption         =   "利息残高表"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   20220
   Icon            =   "RDF011_利息前払未払残高表.dsx":0000
   StartUpPosition =   3  'Windows の既定値
   WindowState     =   2  '最大化
   _ExtentX        =   35666
   _ExtentY        =   19315
   SectionData     =   "RDF011_利息前払未払残高表.dsx":0ECA
End
Attribute VB_Name = "RDF011_利息前払未払残高表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'杉村倉庫仕様
Private Const pPROGRAM_ID As String = "RDF011_利息前払未払残高表"
'
Dim wML As Integer
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
    Dim wWhere As String, wsS As String, wOrder As String
    Dim FLG_Order As Boolean
    
    Dim wstr As String
    Dim wsRet As String
    Dim ws_Ginko As String
    Dim j As Integer, k As Integer, l As Integer, wIndex As Integer
    
    Dim wdate As Date
    Dim w開始年月日 As Date
    Dim w推移表区分 As String, wsNengetu As String
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

    wML = 12 '列数
'
    'パラメータ部分
    Me.PageHeader.Controls("L_金融リストラ番号").Visible = True
    Me.PageHeader.Controls("H00_金融リストラ番号").Visible = True
'
    If GRpt.推移 = "年次" Then
        GRpt.テキスト_01 = GRpt.テキスト_02
        
        If G金利SM = True Then
            L_帳票名.Caption = " " & GRpt.帳票名 & " " & GRpt.テキスト_01 & "年度 - 金利SM " & GRpt.指定 & " " & GRpt.推移 & "- "
        Else
            L_帳票名.Caption = " " & GRpt.帳票名 & " " & GRpt.テキスト_01 & "年度 -" & GRpt.指定 & " " & GRpt.推移 & "- "
        End If
    Else
        GRpt.テキスト_02 = GRpt.テキスト_01
        
        wdate = C年月日.平成To西暦("年月日", GRpt.テキスト_02)
        'GRpt.テキスト_01 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "e")
        '2019/01/15 日付入力区分 仕様変更
        If G基本情報.日付入力区分 = "0" Then
        '和暦
            If Len(GRpt.テキスト_01) <= 2 Then
                GRpt.テキスト_01 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "e")
            Else
                GRpt.テキスト_01 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "y")
            End If
        Else
        '西暦
                GRpt.テキスト_01 = C年月日.年度開始年月日(C年月日.年度変換(CStr(wdate)), "y")
        End If
            
        wsNengetu = GRpt.テキスト_01
        
        If G金利SM = True Then
            L_帳票名.Caption = " " & GRpt.帳票名 & " " & GRpt.テキスト_02 & " - 金利SM " & GRpt.指定 & " " & GRpt.推移 & "- "
        Else
            L_帳票名.Caption = " " & GRpt.帳票名 & " " & GRpt.テキスト_02 & " -" & GRpt.指定 & GRpt.推移 & "- "
        End If
    End If
'
    '----------------------------------------------------------------
    '                           ** 見出し **
    '----------------------------------------------------------------
    出力日 = Now
    企業名 = GCoName
    
    H00_金融リストラ番号 = GRpt.金融
    
    If GRpt.千円単位 = 1 Then
        w分母 = 1000
        L_単位.Caption = "（千円単位）"
    Else
        w分母 = 1
        L_単位 = "（円単位）"
    End If
    
    '----------------------------------------------------------------
    '          ** グループ ・ Where ・ OrderＢｙ **
    '----------------------------------------------------------------
    If GRpt.S_利息 = "分類1" Then
        GroupHeader1.DataField = "GrpFld_Risoku"
    ElseIf GRpt.S_利息 = "分類2" Then
        GroupHeader2.DataField = "GrpFld_Risoku"
    ElseIf GRpt.S_利息 = "分類3" Then
        GroupHeader3.DataField = "GrpFld_Risoku"
    ElseIf GRpt.S_利息 = "分類4" Then
        GroupHeader4.DataField = "GrpFld_Risoku"
    End If
    
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
    If GroupHeader1.DataField = "GrpFld_Risoku" Then
        wsS = wsS & "分類1:利息区分 "
        Me.GroupFooter1.Controls("G10_計名").DataField = "I_利息区分"
        Me.GroupFooter1.Controls("Shape_1").BackColor = &HC0E0FF
    ElseIf GroupHeader1.DataField = "GrpFld_KShubetu" Then
        wsS = wsS & "分類1:借入金種別 "
        Me.GroupFooter1.Controls("G10_計名").DataField = "I_借入金種別名"
        Me.GroupFooter1.Controls("Shape_1").BackColor = &HFFFFC0
    ElseIf GroupHeader1.DataField = "GrpFld_Bumon" Then
        wsS = wsS & "分類1:部門 "
        Me.GroupFooter1.Controls("G10_計名").DataField = "I_部門名"
        Me.GroupFooter1.Controls("Shape_1").BackColor = &HC0FFC0
    ElseIf GroupHeader1.DataField = "GrpFld_Kinyu" Then
        wsS = wsS & "分類1:金融機関 "
        Me.GroupFooter1.Controls("G10_計名").DataField = "I_金融機関名"
        Me.GroupFooter1.Controls("Shape_1").BackColor = &HE0E0E0
    ElseIf GroupHeader1.DataField = "GrpFld_Ginko" Then
        wsS = wsS & "分類1:銀行 "
        Me.GroupFooter1.Controls("G10_計名").DataField = "I_銀行名"
        Me.GroupFooter1.Controls("Shape_1").BackColor = &HC0FFFF
    ElseIf GroupHeader1.DataField = "GrpFld_KGroup" Then
        wsS = wsS & "分類1:金利G "
        Me.GroupFooter1.Controls("G10_計名").DataField = "I_金利グループ名"
        Me.GroupFooter1.Controls("Shape_1").BackColor = C_LGreen
    End If
    
    If GroupHeader2.DataField = "GrpFld_Risoku" Then
        wsS = wsS & "分類2:利息区分 "
        Me.GroupFooter2.Controls("G20_計名").DataField = "I_利息区分"
        Me.GroupFooter2.Controls("Shape_2").BackColor = &HC0E0FF
    ElseIf GroupHeader2.DataField = "GrpFld_KShubetu" Then
        wsS = wsS & "分類2:借入金種別 "
        Me.GroupFooter2.Controls("G20_計名").DataField = "I_借入金種別名"
        Me.GroupFooter2.Controls("Shape_2").BackColor = &HFFFFC0
    ElseIf GroupHeader2.DataField = "GrpFld_Bumon" Then
        wsS = wsS & "分類2:部門 "
        Me.GroupFooter2.Controls("G20_計名").DataField = "I_部門名"
        Me.GroupFooter2.Controls("Shape_2").BackColor = &HC0FFC0
    ElseIf GroupHeader2.DataField = "GrpFld_Kinyu" Then
        wsS = wsS & "分類2:金融機関 "
        Me.GroupFooter2.Controls("G20_計名").DataField = "I_金融機関名"
        Me.GroupFooter2.Controls("Shape_2").BackColor = &HE0E0E0
    ElseIf GroupHeader2.DataField = "GrpFld_Ginko" Then
        wsS = wsS & "分類2:銀行 "
        Me.GroupFooter2.Controls("G20_計名").DataField = "I_銀行名"
        Me.GroupFooter2.Controls("Shape_2").BackColor = &HC0FFFF
    ElseIf GroupHeader2.DataField = "GrpFld_KGroup" Then
        wsS = wsS & "分類2:金利G "
        Me.GroupFooter2.Controls("G20_計名").DataField = "I_金利グループ名"
        Me.GroupFooter2.Controls("Shape_2").BackColor = C_LGreen
    End If
    
    If GroupHeader3.DataField = "GrpFld_Risoku" Then
        wsS = wsS & "分類3:利息区分 "
        Me.GroupFooter3.Controls("G30_計名").DataField = "I_利息区分"
        Me.GroupFooter3.Controls("Shape_3").BackColor = &HC0E0FF
    ElseIf GroupHeader3.DataField = "GrpFld_KShubetu" Then
        wsS = wsS & "分類3:借入金種別 "
        Me.GroupFooter3.Controls("G30_計名").DataField = "I_借入金種別名"
        Me.GroupFooter3.Controls("Shape_3").BackColor = &HFFFFC0
    ElseIf GroupHeader3.DataField = "GrpFld_Bumon" Then
        wsS = wsS & "分類3:部門 "
        Me.GroupFooter3.Controls("G30_計名").DataField = "I_部門名"
        Me.GroupFooter3.Controls("Shape_3").BackColor = &HC0FFC0
    ElseIf GroupHeader3.DataField = "GrpFld_Kinyu" Then
        wsS = wsS & "分類3:金融機関 "
        Me.GroupFooter3.Controls("G30_計名").DataField = "I_金融機関名"
        Me.GroupFooter3.Controls("Shape_3").BackColor = &HE0E0E0
    ElseIf GroupHeader3.DataField = "GrpFld_Ginko" Then
        wsS = wsS & "分類3:銀行 "
        Me.GroupFooter3.Controls("G30_計名").DataField = "I_銀行名"
        Me.GroupFooter3.Controls("Shape_3").BackColor = &HC0FFFF
    ElseIf GroupHeader3.DataField = "GrpFld_KGroup" Then
        wsS = wsS & "分類3:金利G "
        Me.GroupFooter3.Controls("G30_計名").DataField = "I_金利グループ名"
        Me.GroupFooter3.Controls("Shape_3").BackColor = C_LGreen
    End If
    
    If GroupHeader4.DataField = "GrpFld_Risoku" Then
        wsS = wsS & "分類4:利息区分 "
        Me.GroupFooter4.Controls("G40_計名").DataField = "I_利息区分"
        Me.GroupFooter4.Controls("Shape_4").BackColor = &HC0E0FF
    ElseIf GroupHeader4.DataField = "GrpFld_KShubetu" Then
        wsS = wsS & "分類4:借入金種別 "
        Me.GroupFooter4.Controls("G40_計名").DataField = "I_借入金種別名"
        Me.GroupFooter4.Controls("Shape_4").BackColor = &HFFFFC0
    ElseIf GroupHeader4.DataField = "GrpFld_Bumon" Then
        wsS = wsS & "分類4:部門 "
        Me.GroupFooter4.Controls("G40_計名").DataField = "I_部門名"
        Me.GroupFooter4.Controls("Shape_4").BackColor = &HC0FFC0
    ElseIf GroupHeader4.DataField = "GrpFld_Kinyu" Then
        wsS = wsS & "分類4:金融機関 "
        Me.GroupFooter4.Controls("G40_計名").DataField = "I_金融機関名"
        Me.GroupFooter4.Controls("Shape_4").BackColor = &HE0E0E0
    ElseIf GroupHeader4.DataField = "GrpFld_Ginko" Then
        wsS = wsS & "分類4:銀行 "
        Me.GroupFooter4.Controls("G40_計名").DataField = "I_銀行名"
        Me.GroupFooter4.Controls("Shape_4").BackColor = &HC0FFFF
    ElseIf GroupHeader4.DataField = "GrpFld_KGroup" Then
        wsS = wsS & "分類4:金利G "
        Me.GroupFooter4.Controls("G40_計名").DataField = "I_金利グループ名"
        Me.GroupFooter4.Controls("Shape_4").BackColor = C_LGreen
    End If
'
    GroupHeader1.Height = 0
    GroupHeader2.Height = 0
    GroupHeader3.Height = 0
    GroupHeader4.Height = 0
    GroupHeader1.Visible = False
    GroupHeader2.Visible = False
    GroupHeader3.Visible = False
    GroupHeader4.Visible = False
    If GroupHeader1.DataField = "GrpFld_Risoku" Then
        GroupHeader1.Height = 430
        GroupHeader1.Visible = True
    ElseIf GroupHeader2.DataField = "GrpFld_Risoku" Then
        GroupHeader2.Height = 430
        GroupHeader2.Visible = True
    ElseIf GroupHeader3.DataField = "GrpFld_Risoku" Then
        GroupHeader3.Height = 430
        GroupHeader3.Visible = True
    ElseIf GroupHeader4.DataField = "GrpFld_Risoku" Then
        GroupHeader4.Height = 430
        GroupHeader4.Visible = True
    End If
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
    'ReportFooter.Visible = True
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
'
    '改ページ
    Me.GroupFooter1.NewPage = GRpt.NewPage1
    Me.GroupFooter2.NewPage = GRpt.NewPage2
    Me.GroupFooter3.NewPage = GRpt.NewPage3
    Me.GroupFooter4.NewPage = GRpt.NewPage4
    
    '印刷設定
    If GRpt.詳細表示 = 1 Then
        Me.Detail.Height = 210
    Else
        Me.Detail.Height = 0
    End If
    
    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    Call MBD020_借入金ワークテーブル作成(wsTbl) 'データ絞り込み
    Call MRB010_標準入力借入残高表("DCIA010_借入金ワーク")   '16/03/26利子補給に伴う変更
    'Call MRB010_標準入力借入残高表固定日数("DCIA010_借入金ワーク")
    Call MRB010_手入力借入残高表("DCIA010_借入金ワーク")
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

    wIndex = 1
    If GRpt.推移 <> "年次" Then
        For j = 1 To wML
            If GRpt.テキスト_02 = w推移表タイトル.X番目年月(j) Then
                    wIndex = j
                Exit For
            End If
        Next
    End If
    
    'CSVファイル パラメータ:GInt1
    GInt1 = wIndex
'
    '----------------------------------------------------------------
    '                          ** 合計 計算 **
    '----------------------------------------------------------------
'
    '----------------------------------------------------------------
    '                       ** レコード　ソース **
    '----------------------------------------------------------------
    w番号 = Right("00" + CStr(wIndex), 2)
    '** レコード　ソース **
    wstr = "Select "
    wstr = wstr & "K.借入番号 As I_借入番号,"
    
    'セクションGR
    wstr = wstr & "K.利息区分 As GrpFld_Risoku,"
    wstr = wstr & "K.銀行番号 As GrpFld_Ginko,"
    wstr = wstr & "G.金融機関番号 As GrpFld_Kinyu,"
    wstr = wstr & "B.部門番号 As GrpFld_Bumon,"
    wstr = wstr & " K.借入金種別区分 As GrpFld_KShubetu,"
    If GStr = "金利GR" Then
        wstr = wstr & "IIF(K.金利グループ区分<>'',K.金利グループ区分,'99999') As GrpFld_KGroup,"
    End If
    
    wstr = wstr & "IIF(K.利息区分 = '" & XMXA020_区分("利息区分", "利息先払") & "','利息先払','利息後払') As I_利息区分,"
    wstr = wstr & "G.銀行名 As I_銀行名,"
    wstr = wstr & "G.金融機関名 As I_金融機関名,"
    wstr = wstr & "B.部門名 As I_部門名,"
    wstr = wstr & "S.借入金種別名 As I_借入金種別名,"
    If GStr <> "金利GR" Then
        wstr = wstr & "'' As I_金利グループ名,"
    Else
        wstr = wstr & "IIF(KG.金利グループ名<>'',KG.金利グループ名,'グループ無') As I_金利グループ名,"
    End If
    
    wstr = wstr & "K.借入内容 As I_借入内容,"
    wstr = wstr + "IIF(K.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As I_金利種別,"
    wstr = wstr + "KK.基準金利名 As I_基準金利名,"
    wstr = wstr & "K.金利条件 As I_金利備考,"
    'wstr = wstr & "format(K.利率,'#,##0.00000') As I_利率,"
    wstr = wstr & "Format(Z.利率_" & w番号 & ",'#,##0.00000') As I_利率,"
    wstr = wstr & "fix(Z.残高_" & w番号 & " * Z.利率_" & w番号 & "/100) As I_利率融資残高,"
    
    wstr = wstr & "Z.残高_" & w番号 & " As I_融資残高,"
'    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & "-Z.前払利息増_" & w番号 & "+Z.前払利息_" & w番号 & ",-Z.未払利息増_" & w番号 & "+Z.未払利息減_" & w番号 & "+Z.未払利息_" & w番号 & ") As I_前月利息残高,"
'    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息増_" & w番号 & ",Z.未払利息増_" & w番号 & ") As I_利息増,"
'    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & ",Z.未払利息減_" & w番号 & ") As I_利息減,"
    '杉村倉庫仕様
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息増_" & w番号 & ",Z.未払利息減_" & w番号 & ") As I_利息増,"
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & "-Z.前払利息増_" & w番号 & "+Z.前払利息_" & w番号 & ",-Z.未払利息増_" & w番号 & "+Z.未払利息減_" & w番号 & "+Z.未払利息_" & w番号 & ") As I_前月利息残高,"
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息_" & w番号 & ",Z.未払利息_" & w番号 & ") As I_利息残高,"
    wstr = wstr & "IIF(K.利息区分='" & XMXA020_区分("利息区分", "利息先払") & "',Z.前払利息減_" & w番号 & ",Z.未払利息増_" & w番号 & ") As I_利息減,"
    wstr = wstr & "Z2.損益利息額_" & w番号 & " As I_損益利息額"
    
    wstr = wstr & " FROM ((((((DCDA010_借入残高推移表結果 As Z"
    wstr = wstr & " INNER JOIN DCDA010_借入残高推移表結果２ As Z2"
    wstr = wstr & " ON Z.借入番号=Z2.借入番号)"
    wstr = wstr & " INNER JOIN DCIA010_借入金ワーク As K"
    wstr = wstr & " ON Z.借入番号=K.借入番号)"
    wstr = wstr & " INNER JOIN DAAA040_銀行マスタ As G"
    wstr = wstr & " ON K.銀行番号=G.銀行番号)"
    wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
    wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分)"
    wstr = wstr & " LEFT JOIN DAAA200_部門マスタ As B"
    wstr = wstr & " ON K.プロジェクト番号 = B.部門番号)"
'    wstr = wstr & " LEFT JOIN DAAA115_金利シミュレーショングループ As KG"
'    wstr = wstr & " ON K.金利グループ区分 = KG.金利グループ区分)"
    wstr = wstr + " Left Join DAAA116_基準金利 As KK"
    wstr = wstr + " ON K.基準金利区分 = KK.基準金利区分)"
    
    If GStr = "金利GR" Then
        wstr = wstr & " LEFT JOIN DAAA115_金利シミュレーショングループ As KG"
        wstr = wstr & " ON K.金利グループ区分 = KG.金利グループ区分"
    End If
    
    'Where 条件
    'All 0 は表示しない
    wstr = wstr & " Where Z.前払利息増_" & w番号 & "<>0"
    wstr = wstr & " Or Z.前払利息減_" & w番号 & "<>0"
    wstr = wstr & " Or Z.前払利息_" & w番号 & "<>0"
    wstr = wstr & " Or Z.未払利息増_" & w番号 & "<>0"
    wstr = wstr & " Or Z.未払利息減_" & w番号 & "<>0"
    wstr = wstr & " Or Z.未払利息_" & w番号 & "<>0"
    wstr = wstr & " Or Z2.損益利息額_" & w番号 & "<>0"
    
'    If GStr <> "金利GR" Then
'        wstr = wstr & " ORDER BY K.借入金種別区分,K.利息区分,K.銀行番号,K.借入番号"
'    Else
'        '金利SM
'        wstr = wstr & " ORDER BY IIF(K.金利グループ区分<>'',K.金利グループ区分,'99999'),K.利息区分,K.銀行番号,K.借入番号"
'    End If
        
        'Order
'    If GStr <> "金利GR" Then
'        wWhere = wWhere & " ORDER BY K.借入金種別区分,K.銀行番号,Z.借入番号"
'        'wWhere = wWhere & " ORDER By K.銀行番号,借入金種別区分,Z.借入番号"
'    Else
'        '金利SM
'        wWhere = wWhere & " ORDER BY IIF(K.金利グループ区分<>'',K.金利グループ区分,'99999'),K.銀行番号,Z.借入番号"
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
    wOrder = " Order by K.利息区分," & wOrder & "K.借入番号"
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
    'CSVファイル パラメータ:GInt1
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
' Detail_BeforePrint
'------------------------------------------------
Private Sub Detail_BeforePrint()
'
    Me.Detail.Controls("I_融資残高") = Format(P8.FCDbl(Me.Detail.Controls("I_融資残高")) / w分母, "#,##0")
    Me.Detail.Controls("I_前月利息残高") = Format(P8.FCDbl(Me.Detail.Controls("I_前月利息残高")) / w分母, "#,##0")
    Me.Detail.Controls("I_利息増") = Format(P8.FCDbl(Me.Detail.Controls("I_利息増")) / w分母, "#,##0")
    Me.Detail.Controls("I_利息減") = Format(P8.FCDbl(Me.Detail.Controls("I_利息減")) / w分母, "#,##0")
    Me.Detail.Controls("I_利息残高") = Format(P8.FCDbl(Me.Detail.Controls("I_利息残高")) / w分母, "#,##0")
    Me.Detail.Controls("I_損益利息額") = Format(P8.FCDbl(Me.Detail.Controls("I_損益利息額")) / w分母, "#,##0")
'
    Call MXA030_ReportColor(Me.Detail.Controls("I_前月利息残高"))
    Call MXA030_ReportColor(Me.Detail.Controls("I_利息増"))
    Call MXA030_ReportColor(Me.Detail.Controls("I_利息減"))
    Call MXA030_ReportColor(Me.Detail.Controls("I_利息残高"))
    Call MXA030_ReportColor(Me.Detail.Controls("I_損益利息額"))
'
End Sub

'------------------------------------------------
' GroupHeader1_BeforePrint
'------------------------------------------------
Private Sub GroupHeader1_BeforePrint()
'
'    If Me.GroupHeader1.Controls("G10_NAME") = "利息先払" Then
'        Me.GroupHeader1.Controls("G10_NAME") = "前払利息　計上"
'        Me.GroupHeader1.Controls("L10_利息増").Caption = "前払利息増"
'        Me.GroupHeader1.Controls("L10_利息減").Caption = "前払利息減"
'        If GRpt.推移 = "月次" Then
'            Me.GroupHeader1.Controls("L10_利息前残").Caption = "前月前払利息残高"
'            Me.GroupHeader1.Controls("L10_利息残").Caption = "当月前払利息残高"
'        Else
'            Me.GroupHeader1.Controls("L10_利息前残").Caption = "前期前払利息残高"
'            Me.GroupHeader1.Controls("L10_利息残").Caption = "当期前払利息残高"
'        End If
'    ElseIf Me.GroupHeader1.Controls("G10_NAME") = "利息後払" Then
'        Me.GroupHeader1.Controls("G10_NAME") = "未払利息　計上"
'        Me.GroupHeader1.Controls("L10_利息増").Caption = "未払利息増"
'        Me.GroupHeader1.Controls("L10_利息減").Caption = "未払利息減"
'        If GRpt.推移 = "月次" Then
'            Me.GroupHeader1.Controls("L10_利息前残").Caption = "前月未払利息残高"
'            Me.GroupHeader1.Controls("L10_利息残").Caption = "当月未払利息残高"
'        Else
'            Me.GroupHeader1.Controls("L10_利息前残").Caption = "前期未払利息残高"
'            Me.GroupHeader1.Controls("L10_利息残").Caption = "当期未払利息残高"
'        End If
'    End If
    '杉村倉庫仕様
    Me.GroupHeader1.Controls("L10_利息増").Caption = "支払額"
    Me.GroupHeader1.Controls("L10_利息減").Caption = "支払利息"
    If Me.GroupHeader1.Controls("G10_NAME") = "利息先払" Then
        Me.GroupHeader1.Controls("G10_NAME") = "前払利息　計上"
        Me.GroupHeader1.Controls("L10_利息前残").Caption = "前払利息(洗替+)"
        Me.GroupHeader1.Controls("L10_利息残").Caption = "前払利息(計上額-)"
    ElseIf Me.GroupHeader1.Controls("G10_NAME") = "利息後払" Then
        Me.GroupHeader1.Controls("G10_NAME") = "未払利息　計上"
        Me.GroupHeader1.Controls("L10_利息前残").Caption = "未払利息(洗替-)"
        Me.GroupHeader1.Controls("L10_利息残").Caption = "未払利息(計上額+)"
    End If
'
End Sub

'------------------------------------------------
' GroupHeader2_BeforePrint
'------------------------------------------------
Private Sub GroupHeader2_BeforePrint()
'
'    If Me.GroupHeader2.Controls("G20_NAME") = "利息先払" Then
'        Me.GroupHeader2.Controls("G20_NAME") = "前払利息　計上"
'        Me.GroupHeader2.Controls("L20_利息増").Caption = "前払利息増"
'        Me.GroupHeader2.Controls("L20_利息減").Caption = "前払利息減"
'        If GRpt.推移 = "月次" Then
'            Me.GroupHeader2.Controls("L20_利息前残").Caption = "前月前払利息残高"
'            Me.GroupHeader2.Controls("L20_利息残").Caption = "当月前払利息残高"
'        Else
'            Me.GroupHeader2.Controls("L20_利息前残").Caption = "前期前払利息残高"
'            Me.GroupHeader2.Controls("L20_利息残").Caption = "当期前払利息残高"
'        End If
'    ElseIf Me.GroupHeader2.Controls("G20_NAME") = "利息後払" Then
'        Me.GroupHeader2.Controls("G20_NAME") = "未払利息　計上"
'        Me.GroupHeader2.Controls("L20_利息増").Caption = "未払利息増"
'        Me.GroupHeader2.Controls("L20_利息減").Caption = "未払利息減"
'        If GRpt.推移 = "月次" Then
'            Me.GroupHeader2.Controls("L20_利息前残").Caption = "前月未払利息残高"
'            Me.GroupHeader2.Controls("L20_利息残").Caption = "当月未払利息残高"
'        Else
'            Me.GroupHeader2.Controls("L20_利息前残").Caption = "前期未払利息残高"
'            Me.GroupHeader2.Controls("L20_利息残").Caption = "当期未払利息残高"
'        End If
'    End If
    '杉村倉庫仕様
    Me.GroupHeader2.Controls("L20_利息増").Caption = "支払額"
    Me.GroupHeader2.Controls("L20_利息減").Caption = "支払利息"
    If Me.GroupHeader2.Controls("G20_NAME") = "利息先払" Then
        Me.GroupHeader2.Controls("G20_NAME") = "前払利息　計上"
        Me.GroupHeader2.Controls("L20_利息前残").Caption = "前払利息(洗替+)"
        Me.GroupHeader2.Controls("L20_利息残").Caption = "前払利息(計上額-)"
    ElseIf Me.GroupHeader2.Controls("G20_NAME") = "利息後払" Then
        Me.GroupHeader2.Controls("G20_NAME") = "未払利息　計上"
        Me.GroupHeader2.Controls("L20_利息前残").Caption = "未払利息(洗替-)"
        Me.GroupHeader2.Controls("L20_利息残").Caption = "未払利息(計上額+)"
    End If
'
End Sub

'------------------------------------------------
' GroupHeader3_BeforePrint
'------------------------------------------------
Private Sub GroupHeader3_BeforePrint()
'
'    If Me.GroupHeader3.Controls("G30_NAME") = "利息先払" Then
'        Me.GroupHeader3.Controls("G30_NAME") = "前払利息　計上"
'        Me.GroupHeader3.Controls("L30_利息増").Caption = "前払利息増"
'        Me.GroupHeader3.Controls("L30_利息減").Caption = "前払利息減"
'        If GRpt.推移 = "月次" Then
'            Me.GroupHeader3.Controls("L30_利息前残").Caption = "前月前払利息残高"
'            Me.GroupHeader3.Controls("L30_利息残").Caption = "当月前払利息残高"
'        Else
'            Me.GroupHeader3.Controls("L30_利息前残").Caption = "前期前払利息残高"
'            Me.GroupHeader3.Controls("L30_利息残").Caption = "当期前払利息残高"
'        End If
'    ElseIf Me.GroupHeader3.Controls("G30_NAME") = "利息後払" Then
'        Me.GroupHeader3.Controls("G30_NAME") = "未払利息　計上"
'        Me.GroupHeader3.Controls("L30_利息増").Caption = "未払利息増"
'        Me.GroupHeader3.Controls("L30_利息減").Caption = "未払利息減"
'        If GRpt.推移 = "月次" Then
'            Me.GroupHeader3.Controls("L30_利息前残").Caption = "前月未払利息残高"
'            Me.GroupHeader3.Controls("L30_利息残").Caption = "当月未払利息残高"
'        Else
'            Me.GroupHeader3.Controls("L30_利息前残").Caption = "前期未払利息残高"
'            Me.GroupHeader3.Controls("L30_利息残").Caption = "当期未払利息残高"
'        End If
'    End If
    '杉村倉庫仕様
    Me.GroupHeader3.Controls("L30_利息増").Caption = "支払額"
    Me.GroupHeader3.Controls("L30_利息減").Caption = "支払利息"
    If Me.GroupHeader3.Controls("G30_NAME") = "利息先払" Then
        Me.GroupHeader3.Controls("G30_NAME") = "前払利息　計上"
        Me.GroupHeader3.Controls("L30_利息前残").Caption = "前払利息(洗替+)"
        Me.GroupHeader3.Controls("L30_利息残").Caption = "前払利息(計上額-)"
    ElseIf Me.GroupHeader3.Controls("G30_NAME") = "利息後払" Then
        Me.GroupHeader3.Controls("G30_NAME") = "未払利息　計上"
        Me.GroupHeader3.Controls("L30_利息前残").Caption = "未払利息(洗替-)"
        Me.GroupHeader3.Controls("L30_利息残").Caption = "未払利息(計上額+)"
    End If
'
End Sub

'------------------------------------------------
' GroupHeader4_BeforePrint
'------------------------------------------------
Private Sub GroupHeader4_BeforePrint()
'
'    If Me.GroupHeader4.Controls("G40_NAME") = "利息先払" Then
'        Me.GroupHeader4.Controls("G40_NAME") = "前払利息　計上"
'        Me.GroupHeader4.Controls("L40_利息増").Caption = "前払利息増"
'        Me.GroupHeader4.Controls("L40_利息減").Caption = "前払利息減"
'        If GRpt.推移 = "月次" Then
'            Me.GroupHeader4.Controls("L40_利息前残").Caption = "前月前払利息残高"
'            Me.GroupHeader4.Controls("L40_利息残").Caption = "当月前払利息残高"
'        Else
'            Me.GroupHeader4.Controls("L40_利息前残").Caption = "前期前払利息残高"
'            Me.GroupHeader4.Controls("L40_利息残").Caption = "当期前払利息残高"
'        End If
'    ElseIf Me.GroupHeader4.Controls("G40_NAME") = "利息後払" Then
'        Me.GroupHeader4.Controls("G40_NAME") = "未払利息　計上"
'        Me.GroupHeader4.Controls("L40_利息増").Caption = "未払利息増"
'        Me.GroupHeader4.Controls("L40_利息減").Caption = "未払利息減"
'        If GRpt.推移 = "月次" Then
'            Me.GroupHeader4.Controls("L40_利息前残").Caption = "前月未払利息残高"
'            Me.GroupHeader4.Controls("L40_利息残").Caption = "当月未払利息残高"
'        Else
'            Me.GroupHeader4.Controls("L40_利息前残").Caption = "前期未払利息残高"
'            Me.GroupHeader4.Controls("L40_利息残").Caption = "当期未払利息残高"
'        End If
'    End If
    '杉村倉庫仕様
    Me.GroupHeader4.Controls("L40_利息増").Caption = "支払額"
    Me.GroupHeader4.Controls("L40_利息減").Caption = "支払利息"
    If Me.GroupHeader4.Controls("G40_NAME") = "利息先払" Then
        Me.GroupHeader4.Controls("G40_NAME") = "前払利息　計上"
        Me.GroupHeader4.Controls("L40_利息前残").Caption = "前払利息(洗替+)"
        Me.GroupHeader4.Controls("L40_利息残").Caption = "前払利息(計上額-)"
    ElseIf Me.GroupHeader4.Controls("G40_NAME") = "利息後払" Then
        Me.GroupHeader4.Controls("G40_NAME") = "未払利息　計上"
        Me.GroupHeader4.Controls("L40_利息前残").Caption = "未払利息(洗替-)"
        Me.GroupHeader4.Controls("L40_利息残").Caption = "未払利息(計上額+)"
    End If
'
End Sub

'------------------------------------------------
' GroupFooter1_BeforePrint
'------------------------------------------------
Private Sub GroupFooter1_BeforePrint()
'
    If GRpt.S_利息 = "分類1" Then
        If Me.GroupFooter1.Controls("G10_計名") = "利息先払" Then
            Me.GroupFooter1.Controls("G10_計名") = "前払費用　計"
        ElseIf Me.GroupFooter1.Controls("G10_計名") = "利息後払" Then
            Me.GroupFooter1.Controls("G10_計名") = "未払費用　計"
        End If
    Else
        Me.GroupFooter1.Controls("G10_計名") = Me.GroupFooter1.Controls("G10_計名") & "　計"
    End If
'
    Me.GroupFooter1.Controls("G10_融資残高") = Format(P8.FCDbl(Me.GroupFooter1.Controls("G10_融資残高")) / w分母, "#,##0")
    Me.GroupFooter1.Controls("G10_前月利息残高") = Format(P8.FCDbl(Me.GroupFooter1.Controls("G10_前月利息残高")) / w分母, "#,##0")
    Me.GroupFooter1.Controls("G10_利息増") = Format(P8.FCDbl(Me.GroupFooter1.Controls("G10_利息増")) / w分母, "#,##0")
    Me.GroupFooter1.Controls("G10_利息減") = Format(P8.FCDbl(Me.GroupFooter1.Controls("G10_利息減")) / w分母, "#,##0")
    Me.GroupFooter1.Controls("G10_利息残高") = Format(P8.FCDbl(Me.GroupFooter1.Controls("G10_利息残高")) / w分母, "#,##0")
    Me.GroupFooter1.Controls("G10_損益利息額") = Format(P8.FCDbl(Me.GroupFooter1.Controls("G10_損益利息額")) / w分母, "#,##0")
'
    Call MXA030_ReportColor(Me.GroupFooter1.Controls("G10_前月利息残高"))
    Call MXA030_ReportColor(Me.GroupFooter1.Controls("G10_利息増"))
    Call MXA030_ReportColor(Me.GroupFooter1.Controls("G10_利息減"))
    Call MXA030_ReportColor(Me.GroupFooter1.Controls("G10_利息残高"))
    Call MXA030_ReportColor(Me.GroupFooter1.Controls("G10_損益利息額"))
'
End Sub

'------------------------------------------------
' GroupFooter2_BeforePrint
'------------------------------------------------
Private Sub GroupFooter2_BeforePrint()
'
    If GRpt.S_利息 = "分類2" Then
        If Me.GroupFooter2.Controls("G20_計名") = "利息先払" Then
            Me.GroupFooter2.Controls("G20_計名") = "前払費用　計"
        ElseIf Me.GroupFooter2.Controls("G20_計名") = "利息後払" Then
            Me.GroupFooter2.Controls("G20_計名") = "未払費用　計"
        End If
    Else
        Me.GroupFooter2.Controls("G20_計名") = Me.GroupFooter2.Controls("G20_計名") & "　計"
    End If
'
    Me.GroupFooter2.Controls("G20_融資残高") = Format(P8.FCDbl(Me.GroupFooter2.Controls("G20_融資残高")) / w分母, "#,##0")
    Me.GroupFooter2.Controls("G20_前月利息残高") = Format(P8.FCDbl(Me.GroupFooter2.Controls("G20_前月利息残高")) / w分母, "#,##0")
    Me.GroupFooter2.Controls("G20_利息増") = Format(P8.FCDbl(Me.GroupFooter2.Controls("G20_利息増")) / w分母, "#,##0")
    Me.GroupFooter2.Controls("G20_利息減") = Format(P8.FCDbl(Me.GroupFooter2.Controls("G20_利息減")) / w分母, "#,##0")
    Me.GroupFooter2.Controls("G20_利息残高") = Format(P8.FCDbl(Me.GroupFooter2.Controls("G20_利息残高")) / w分母, "#,##0")
    Me.GroupFooter2.Controls("G20_損益利息額") = Format(P8.FCDbl(Me.GroupFooter2.Controls("G20_損益利息額")) / w分母, "#,##0")
'
    Call MXA030_ReportColor(Me.GroupFooter2.Controls("G20_前月利息残高"))
    Call MXA030_ReportColor(Me.GroupFooter2.Controls("G20_利息増"))
    Call MXA030_ReportColor(Me.GroupFooter2.Controls("G20_利息減"))
    Call MXA030_ReportColor(Me.GroupFooter2.Controls("G20_利息残高"))
    Call MXA030_ReportColor(Me.GroupFooter2.Controls("G20_損益利息額"))
'
End Sub

'------------------------------------------------
' GroupFooter3_BeforePrint
'------------------------------------------------
Private Sub GroupFooter3_BeforePrint()
'
    If GRpt.S_利息 = "分類3" Then
        If Me.GroupFooter3.Controls("G30_計名") = "利息先払" Then
            Me.GroupFooter3.Controls("G30_計名") = "前払費用　計"
        ElseIf Me.GroupFooter3.Controls("G30_計名") = "利息後払" Then
            Me.GroupFooter3.Controls("G30_計名") = "未払費用　計"
        End If
    Else
        Me.GroupFooter3.Controls("G30_計名") = Me.GroupFooter3.Controls("G30_計名") & "　計"
    End If
'
    Me.GroupFooter3.Controls("G30_融資残高") = Format(P8.FCDbl(Me.GroupFooter3.Controls("G30_融資残高")) / w分母, "#,##0")
    Me.GroupFooter3.Controls("G30_前月利息残高") = Format(P8.FCDbl(Me.GroupFooter3.Controls("G30_前月利息残高")) / w分母, "#,##0")
    Me.GroupFooter3.Controls("G30_利息増") = Format(P8.FCDbl(Me.GroupFooter3.Controls("G30_利息増")) / w分母, "#,##0")
    Me.GroupFooter3.Controls("G30_利息減") = Format(P8.FCDbl(Me.GroupFooter3.Controls("G30_利息減")) / w分母, "#,##0")
    Me.GroupFooter3.Controls("G30_利息残高") = Format(P8.FCDbl(Me.GroupFooter3.Controls("G30_利息残高")) / w分母, "#,##0")
    Me.GroupFooter3.Controls("G30_損益利息額") = Format(P8.FCDbl(Me.GroupFooter3.Controls("G30_損益利息額")) / w分母, "#,##0")
'
    Call MXA030_ReportColor(Me.GroupFooter3.Controls("G30_前月利息残高"))
    Call MXA030_ReportColor(Me.GroupFooter3.Controls("G30_利息増"))
    Call MXA030_ReportColor(Me.GroupFooter3.Controls("G30_利息減"))
    Call MXA030_ReportColor(Me.GroupFooter3.Controls("G30_利息残高"))
    Call MXA030_ReportColor(Me.GroupFooter3.Controls("G30_損益利息額"))
'
End Sub

'------------------------------------------------
' GroupFooter4_BeforePrint
'------------------------------------------------
Private Sub GroupFooter4_BeforePrint()
'
    If GRpt.S_利息 = "分類4" Then
        If Me.GroupFooter4.Controls("G40_計名") = "利息先払" Then
            Me.GroupFooter4.Controls("G40_計名") = "前払費用　計"
        ElseIf Me.GroupFooter4.Controls("G40_計名") = "利息後払" Then
            Me.GroupFooter4.Controls("G40_計名") = "未払費用　計"
        End If
    Else
        Me.GroupFooter4.Controls("G40_計名") = Me.GroupFooter4.Controls("G40_計名") & "　計"
    End If
'
    Me.GroupFooter4.Controls("G40_融資残高") = Format(P8.FCDbl(Me.GroupFooter4.Controls("G40_融資残高")) / w分母, "#,##0")
    Me.GroupFooter4.Controls("G40_前月利息残高") = Format(P8.FCDbl(Me.GroupFooter4.Controls("G40_前月利息残高")) / w分母, "#,##0")
    Me.GroupFooter4.Controls("G40_利息増") = Format(P8.FCDbl(Me.GroupFooter4.Controls("G40_利息増")) / w分母, "#,##0")
    Me.GroupFooter4.Controls("G40_利息減") = Format(P8.FCDbl(Me.GroupFooter4.Controls("G40_利息減")) / w分母, "#,##0")
    Me.GroupFooter4.Controls("G40_利息残高") = Format(P8.FCDbl(Me.GroupFooter4.Controls("G40_利息残高")) / w分母, "#,##0")
    Me.GroupFooter4.Controls("G40_損益利息額") = Format(P8.FCDbl(Me.GroupFooter4.Controls("G40_損益利息額")) / w分母, "#,##0")
'
    Call MXA030_ReportColor(Me.GroupFooter4.Controls("G40_前月利息残高"))
    Call MXA030_ReportColor(Me.GroupFooter4.Controls("G40_利息増"))
    Call MXA030_ReportColor(Me.GroupFooter4.Controls("G40_利息減"))
    Call MXA030_ReportColor(Me.GroupFooter4.Controls("G40_利息残高"))
    Call MXA030_ReportColor(Me.GroupFooter4.Controls("G40_損益利息額"))
'
End Sub

'------------------------------------------------
' ReportFooter_BeforePrint
'------------------------------------------------
Private Sub ReportFooter_BeforePrint()
'
    Me.ReportFooter.Controls("G90_融資残高") = Format(P8.FCDbl(Me.ReportFooter.Controls("G90_融資残高")) / w分母, "#,##0")
    Me.ReportFooter.Controls("G90_前月利息残高") = Format(P8.FCDbl(Me.ReportFooter.Controls("G90_前月利息残高")) / w分母, "#,##0")
    Me.ReportFooter.Controls("G90_利息増") = Format(P8.FCDbl(Me.ReportFooter.Controls("G90_利息増")) / w分母, "#,##0")
    Me.ReportFooter.Controls("G90_利息減") = Format(P8.FCDbl(Me.ReportFooter.Controls("G90_利息減")) / w分母, "#,##0")
    Me.ReportFooter.Controls("G90_利息残高") = Format(P8.FCDbl(Me.ReportFooter.Controls("G90_利息残高")) / w分母, "#,##0")
    Me.ReportFooter.Controls("G90_損益利息額") = Format(P8.FCDbl(Me.ReportFooter.Controls("G90_損益利息額")) / w分母, "#,##0")
'
    Call MXA030_ReportColor(Me.ReportFooter.Controls("G90_前月利息残高"))
    Call MXA030_ReportColor(Me.ReportFooter.Controls("G90_利息増"))
    Call MXA030_ReportColor(Me.ReportFooter.Controls("G90_利息減"))
    Call MXA030_ReportColor(Me.ReportFooter.Controls("G90_前月利息残高"))
    Call MXA030_ReportColor(Me.ReportFooter.Controls("G90_損益利息額"))
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

