VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RDA020_借入一覧表 
   Caption         =   "借入一覧表"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10380
   Icon            =   "RDA020_借入一覧表.dsx":0000
   StartUpPosition =   2  '画面の中央
   WindowState     =   2  '最大化
   _ExtentX        =   18309
   _ExtentY        =   11933
   SectionData     =   "RDA020_借入一覧表.dsx":0ECA
End
Attribute VB_Name = "RDA020_借入一覧表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Const pPROGRAM_ID As String = "RDA020_借入一覧表"

Dim wstr As String, wstr2 As String, wStr3 As String
Dim wRs As Recordset, wRs2 As Recordset, wRs3 As Recordset

Dim wsTbl As String

Dim w残高年月 As Date, w残高年月月始 As Date
Dim w残高年月min As Date            '5/9/6 V129
Dim w管理年月日 As Date             '5/9/6 V129

Dim w年 As Integer, w月 As Integer, w日 As Integer
Dim w分母 As Integer

Dim wyymm As Long
Dim w銀行マスタ As MAA030_銀行
Dim w金融リストラ As String
Dim w借入番号 As String
Dim w借入計画番号 As String
Dim w借入金 As MAA910_借入金

Dim FLG_Pj As Boolean
'
'------------------------------------------------
' ActiveReport_ReportStart
'------------------------------------------------
Private Sub ActiveReport_ReportStart()
'
    Dim j As Integer
    Dim wOrder As String
    Dim FLG_Order As Boolean
    
    Dim w期首借入年度 As String                 '5/10/9 V129 支店への貸付　リストラ番号
    Dim w支店貸付 As String                     '5/10/9 V129 支店への貸付　リストラ番号
    Dim w全社借入 As String                     '5/10/9 V129 支店への貸付　リストラ番号
    Dim w融資残高 As Double                     ' 07/02/09 V180
    Dim w借入金管理区分 As String               ' 07/02/18 V180
    
    Dim FLG_Data As Boolean
    Dim ws01 As String, ws02 As String, wsS As String
    
    '17/10/25 利子補給に伴う変更
    Dim wiFlgRishiHokyu As Integer
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
'
    '帳票名
    'Me.PageHeader.Controls("L_帳票名").Left = 6803

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
    
    'タイトル
    'Me.PageHeader.Controls("L_借入計画番号").Visible = True
    'Me.PageHeader.Controls("Line41").Visible = True
    
    '
    'Me.Detail.Controls("I_借入計画番号").Visible = True
    'Me.Detail.Controls("Line21").Visible = True
    'Me.Detail.Controls("I_借入内容").Width = 3898
    
'    If GProduct <> "金剛石" Then
'        '帳票名
'        Me.PageHeader.Controls("L_帳票名").Left = 4819
'
'        'パラメータ部分
'        Me.PageHeader.Controls("L_計画番号").Visible = False
'        Me.PageHeader.Controls("L_最終実績年月").Visible = False
'
'        Me.PageHeader.Controls("H00_借入計画番号").Visible = False
'        Me.PageHeader.Controls("H00_最終実績年月").Visible = False
'
'        Me.PageHeader.Controls("Line5").Visible = False
'        Me.PageHeader.Controls("Line6").Visible = False
'
'        If GRpt.金融 = "" Then
'            Me.PageHeader.Controls("L_金融リストラ番号").Visible = False
'            Me.PageHeader.Controls("H00_金融リストラ番号").Visible = False
'
'            Me.PageHeader.Controls("Shape1").Visible = False
'            Me.PageHeader.Controls("Shape2").Visible = False
'            Me.PageHeader.Controls("Shape3").Visible = False
'            Me.PageHeader.Controls("Line4").Visible = False
'        Else
'            Me.PageHeader.Controls("L_金融リストラ番号").Left = 0
'            Me.PageHeader.Controls("H00_金融リストラ番号").Left = 0
'
'            Me.PageHeader.Controls("Shape1").Width = 2268
'            Me.PageHeader.Controls("Shape2").Width = 2268
'            Me.PageHeader.Controls("Shape3").Width = 2268
'            Me.PageHeader.Controls("Line4").X2 = 2268
'        End If
'
'        'タイトル
'        Me.PageHeader.Controls("L_借入計画番号").Visible = False
'        Me.PageHeader.Controls("Line41").Visible = False
'
'        '
'        Me.Detail.Controls("I_借入計画番号").Visible = False
'        Me.Detail.Controls("Line21").Visible = False
'        Me.Detail.Controls("I_借入内容").Width = 6166
'    End If
'
    '06/02/01 V150
    Select Case GRpt.帳票名
    Case "借入一覧表"
        wsTbl = "DBDA010_借入金"
        
        Me.PageHeader.Controls("L_番号").Caption = "借入番号"
        Me.PageHeader.Controls("L_内容").Caption = "借入内容"
        'Me.PageHeader.Controls("L_計画番号").Caption = "借入計画番号"
    Case "貸付一覧表"
        wsTbl = "DBDA010_貸付金"
    
        Me.PageHeader.Controls("L_番号").Caption = "貸付番号"
        Me.PageHeader.Controls("L_内容").Caption = "貸付内容"
        'Me.PageHeader.Controls("L_計画番号").Caption = "貸付計画番号"
    End Select
'
    '----------------------------------------------------------------
    '                           ** 見出し **
    '----------------------------------------------------------------
    出力日 = Now
    企業名 = GCoName
    
    'H00_最終実績年月 = ""
    'If Gコントロール.最終実績年月 > CDate("2001/01/01") Then
    '    H00_最終実績年月 = Format(Gコントロール.最終実績年月, Gfmt年月)
    'End If
    
    'H00_借入計画番号 = GRpt.コンボ_02
    H00_金融リストラ番号 = GRpt.金融
    
    If GRpt.千円単位 = 1 Then
        w分母 = 1000
        L_単位.Caption = "（千円単位）"
    Else
        w分母 = 1
        L_単位 = "（円単位）"
    End If
    
    L_帳票名.Caption = " " + GRpt.帳票名 + " - " & GRpt.テキスト_01 & " -"
    'If GRpt.借入金管理区分 = P8.FCDbl(XMXA020_区分("借入金管理区分", "決算用")) Then
    '    L_帳票名.Caption = " " + GRpt.帳票名 + " - " & GRpt.テキスト_01 & " 決算用 -"
    'Else
    '    L_帳票名.Caption = " " + GRpt.帳票名 + " - " & GRpt.テキスト_01 & " 管理用 -"
    'End If
    
    '----------------------------------------------------------------
    '          ** グループ ・ Where ・ OrderＢｙ **
    '----------------------------------------------------------------
'    GRet = MAA080_プロジェクト件数
'    FLG_Pj = False
'    If GRet > 0 Then
'        FLG_Pj = True
'    End If
'
'    'グループセット
'    GroupHeader1.DataField = "GrpFld_Ginko"
'
'    'プロジェクトを指定しない時は表示しない
'
    If GRpt.S_種別 = "分類1" Then
        GroupHeader1.DataField = "GrpFld_KShubetu"
    ElseIf GRpt.S_種別 = "分類2" Then
        GroupHeader2.DataField = "GrpFld_KShubetu"
    End If
    
    If GRpt.S_部門 = "分類1" Then
        GroupHeader1.DataField = "GrpFld_Bumon"
    ElseIf GRpt.S_部門 = "分類2" Then
        GroupHeader2.DataField = "GrpFld_Bumon"
    End If
    
    If GRpt.S_金融 = "分類1" Then
        GroupHeader1.DataField = "GrpFld_Kinyu"
    ElseIf GRpt.S_金融 = "分類2" Then
        GroupHeader2.DataField = "GrpFld_Kinyu"
    End If
    
    If GRpt.S_銀行 = "分類1" Then
        GroupHeader1.DataField = "GrpFld_Ginko"
    ElseIf GRpt.S_銀行 = "分類2" Then
        GroupHeader2.DataField = "GrpFld_Ginko"
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
    End If
    
    If GroupHeader2.DataField = "GrpFld_KShubetu" Then
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
    End If
'
    GroupFooter1.Visible = True
    GroupFooter2.Visible = True
    If GroupHeader1.DataField = "" Then
        GroupFooter1.Visible = False
    End If
    If GroupHeader2.DataField = "" Then
        GroupFooter2.Visible = False
    End If
'
    '帳票指示
    Me.PageHeader.Controls("L_帳票指示").Caption = wsS
'
    '改ページ
    Me.GroupFooter1.NewPage = GRpt.NewPage1
    Me.GroupFooter2.NewPage = GRpt.NewPage2
'
    '----------------------------------------------------------------
    '                       ** 印字用ファイル作成 **
    '----------------------------------------------------------------
    w金融リストラ = GRpt.金融
    w借入計画番号 = GRpt.借入
    
    GVar1 = C年月日.平成To西暦("年月", GRpt.テキスト_01)
    w残高年月 = MBA010_締日年月日(CDate(GVar1))         '2010/08/30
    'w残高年月 = CDate(GVar1)
    'If G基本情報.決算締日 = 31 Then                     '5/9/2 V129
    '    w残高年月 = C年月日.GetDate("月末", w残高年月)
    '    w年 = Year(w残高年月)                           '5/9/2 V129
    '    w月 = Month(w残高年月)                          '5/9/2 V129
    '    w日 = Day(w残高年月)                            '5/9/2 V129
    '    w残高年月min = CDate(CStr(w年) + "/" + CStr(w月) + "/" + CStr(1))  '5/9/2 V129
    '    w残高年月min = DateAdd("d", -1, w残高年月min)          '5/9/6 V129
    '
    'Else                                                '5/9/2 V129
    '    w年 = Year(w残高年月)                           '5/9/2 V129
    '    w月 = Month(w残高年月)                          '5/9/2 V129
    '    w日 = Day(w残高年月)                            '5/9/2 V129
    '    w残高年月 = CDate(CStr(w年) + "/" + CStr(w月) + "/" + CStr(G基本情報.決算締日))  '5/9/2 V129
    '    w残高年月min = DateAdd("m", -1, w残高年月)      '5/9/6 V129
    'End If                                              '5/9/2 V129
    
    If w借入計画番号 = "" Then                          '5/10/17 V129
        w期首借入年度 = C年月日.年度開始年月日(CStr(w年), "e")      '5/10/17 V129
    Else                                               '5/10/17 V129
        w期首借入年度 = Left$(w借入計画番号, 2)          '5/10/9 V129
    End If                                             '5/10/17 V129
    
    w支店貸付 = "支店貸付"                              '5/10/9 V129
    w全社借入 = "全社借入"                              '5/10/17 V129
'
    '** ワークファイル 削除 **
    wstr = ""
    wstr = wstr & "Delete * From DCDA030_借入一覧表"
    GDb.Execute wstr
'
    wstr2 = ""
    wstr2 = wstr2 + "Select * From DCDA030_借入一覧表"
    Call AdoRecordsetOpen(GDb, wRs2, wstr2)
    
        wstr = ""
        wstr = wstr & "Select * From " & wsTbl      '06/02/01 V150
        
        '実データ
        wstr = wstr & " Where ((Format(実行日,'yyyymmdd') <= '" & Format(w残高年月, "yyyymmdd") & "'"
        wstr = wstr & " and Format(最終返済実行日,'yyyymmdd') > '" & Format(w残高年月, "yyyymmdd") & "')"
        wstr = wstr & " and sm区分 = 0"
        wstr = wstr & " and 手入力区分 <> 2 And 取消フラグ = 0)"
        
        If w金融リストラ <> "" Then
        'smデータ
            '2
            wstr = wstr & " Or ((Format(実行日,'yyyymmdd') <= '" & Format(w残高年月, "yyyymmdd") & "'"
            wstr = wstr & " and Format(最終返済実行日,'yyyymmdd') > '" & Format(w残高年月, "yyyymmdd") & "')"
            wstr = wstr & " and sm区分 = 1"
            If w借入計画番号 <> "" Then
                wstr = wstr & " and 借入計画番号 = '" & w借入計画番号 & "'"
            End If
            If w金融リストラ <> "" Then
                wstr = wstr & " and 金融リストラ番号 = '" & w金融リストラ & "'"
            End If
            wstr = wstr & " and 手入力区分 <> 2 And 取消フラグ = 0)"
            
            '3
            If w金融リストラ = "" And w金融リストラ <> w期首借入年度 Then
                wstr = wstr & " Or ((Format(実行日,'yyyymmdd') <= '" & Format(w残高年月, "yyyymmdd") & "'"
                wstr = wstr & " and Format(最終返済実行日,'yyyymmdd') > '" & Format(w残高年月, "yyyymmdd") & "')"
                wstr = wstr & " and sm区分 = 1"
                If w借入計画番号 <> "" Then
                    wstr = wstr & " and 借入計画番号 = '" & w借入計画番号 & "'"
                End If
                wstr = wstr & " and 金融リストラ番号 = '" & w期首借入年度 & "' And 手入力区分 <> 2 And 取消フラグ = 0)" '5/10/9 V129
            End If
            
            '4
            If w金融リストラ = "" And w金融リストラ <> w支店貸付 Then
                wstr = wstr & " Or ((Format(実行日,'yyyymmdd') <= '" & Format(w残高年月, "yyyymmdd") & "'"
                wstr = wstr & " and Format(最終返済実行日,'yyyymmdd') > '" & Format(w残高年月, "yyyymmdd") & "')"
                wstr = wstr & " and sm区分 = 1"
                If w借入計画番号 <> "" Then
                    wstr = wstr & " and 借入計画番号 = '" & w借入計画番号 & "'"
                End If
                wstr = wstr & " and 金融リストラ番号 = '" & w支店貸付 & "' And 手入力区分 <> 2 And 取消フラグ = 0)" '5/10/9 V129
            End If
            
            '5
            If w金融リストラ = "" And w金融リストラ <> w全社借入 Then
                wstr = wstr & " Or ((Format(実行日,'yyyymmdd') <= '" & Format(w残高年月, "yyyymmdd") & "'"
                wstr = wstr & " and Format(最終返済実行日,'yyyymmdd') > '" & Format(w残高年月, "yyyymmdd") & "')"
                wstr = wstr & " and sm区分 = 1"
                If w借入計画番号 <> "" Then
                    wstr = wstr & " and 借入計画番号 = '" & w借入計画番号 & "'"
                End If
                wstr = wstr & " and 金融リストラ番号 = '" & w全社借入 & "' And 手入力区分 <> 2 And 取消フラグ = 0)" '5/10/9 V129
            End If
        End If
        Call AdoRecordsetOpen(GDb, wRs, wstr)
            Do Until wRs.EOF
                                
            w借入金 = MBD010_借入データセット(wRs)
            
            '解約Check
            FLG_Data = True
            If P8.FCStr(w借入金.解約実行日) <> "" And Format(w借入金.解約実行日, "yyyymmdd") < Format(w残高年月, "yyyymmdd") Then
                FLG_Data = False
            End If
            
            If (w借入金.金融リストラ番号 <> "" And w借入金.金融リストラ番号 = w金融リストラ) _
            And P8.FCStr(w借入金.金融解約実行日) <> "" And Format(w借入金.金融解約実行日, "yyyymmdd") < Format(w残高年月, "yyyymmdd") Then
                FLG_Data = False
            End If
            
            If FLG_Data = True Then
                
                w融資残高 = 0
                If w借入金.手入力区分 = P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then                          ' 07/02/09 V180
         
                    '** 借入金テーブル セット **
                    Call MBD010_借入金テーブル作成(w金融リストラ, w借入金)
                    If GRpt.借入金管理区分 = P8.FCDbl(XMXA020_区分("借入金管理区分", "決算用")) Then                       ' 07/02/18 V180
                        w借入金管理区分 = "1"                           ' 07/02/18 V180
                    Else                                                ' 07/02/18 V180
                        w借入金管理区分 = "0"                           ' 07/02/18 V180
                    End If                                              ' 07/02/18 V180
                    w融資残高 = MBD010_借入金標準入力残高(w借入金, w金融リストラ, 1, w残高年月, w借入金管理区分)  ' 07/02/26 V180
                    'w融資残高 = MBD010_借入指定年月融資残高(w借入金, CDate(w残高年月))
                
                Else                                                    ' 07/02/18 V180

                    Call MBD010_借入金入力明細Read(w借入金) 'V180
                
                    w融資残高 = MBD010_借入金手入力残高(w借入金, 1, w残高年月) ' 07/02/09 V180
                    
                    'w借入金.利率 = 0
                    'w借入金.保証料率 = 0
                    w借入金.返済単位月数 = 0
                    'w借入金.金利種別 = 0
                    
                    w借入金.毎月返済額 = 0
                    w借入金.初回返済額 = 0
                    w借入金.最終返済額 = 0

                End If                                                  ' 07/02/18 V180
                
                wRs2.AddNew
                                                
                            'wRs2("入力借入計画番号") = GRpt.コンボ_02
                            wRs2("入力借入計画番号") = w借入金.借入計画番号
                            
                            wRs2("最終実績") = Format(Gコントロール.最終実績年月, "yyyy/mm/dd")
                            wRs2("入力金融リストラ番号") = GRpt.金融
                            wRs2("借入番号") = w借入金.借入番号
                            wRs2("借入内容") = w借入金.借入内容
                            wRs2("銀行番号") = w借入金.銀行番号
                            
                            If GRpt.金融 = "" Or GRpt.金融 = w借入金.金融リストラ番号 Then        'V120
                                wRs2("金融リストラ番号") = w借入金.金融リストラ番号  'V120
                            Else                                                     'V120
                                wRs2("金融リストラ番号") = " "                       'V120
                            End If                                                   'V120
                            
                            wRs2("融資金額") = w借入金.融資金額
                            wRs2("入力残高年月") = w残高年月
                            
                            wRs2("実行日") = w借入金.実行日
                            wRs2("最終返済実行日") = w借入金.最終返済実行日
                            wRs2("初回返済年月") = w借入金.初回返済年月
                            wRs2("最終返済年月") = w借入金.最終返済年月
                            wRs2("解約実行日") = w借入金.解約実行日
                            wRs2("金融解約実行日") = w借入金.金融解約実行日
                            wRs2("毎月返済額") = w借入金.毎月返済額
                            wRs2("初回返済額") = w借入金.初回返済額
                            wRs2("最終返済額") = w借入金.最終返済額
                            wRs2("融資残高") = w融資残高
                            
                            '変動金利の場合
                            If P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) = w借入金.金利種別 Then
                                If w借入金.変動最終利率 > -1 Then
                                    wRs2("利率") = w借入金.変動最終利率
                                    wRs2("変動利率フラグ") = 1
                                Else
                                    wRs2("利率") = w借入金.利率
                                    wRs2("変動利率フラグ") = 0
                                End If
                            Else
                                wRs2("利率") = w借入金.利率
                                wRs2("変動利率フラグ") = 0
                            End If
                            
                            wRs2("保証料率") = w借入金.保証料率
                            wRs2("返済単位月数") = w借入金.返済単位月数
                            wRs2("有担保フラグ") = w借入金.有担保フラグ
                            wRs2("金利種別") = w借入金.金利種別
                            wRs2("金利条件") = w借入金.金利条件
                            
                            '手入力の場合
                            If w借入金.手入力区分 <> P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then
                                'wRs2("利率") = Null
                                'wRs2("変動利率フラグ") = Null
                                'wRs2("保証料率") = Null
                                wRs2("返済単位月数") = Null
                                'wRs2("有担保フラグ") = Null
                                'wRs2("金利種別") = Null
                                'wRs2("金利条件") = Null
                            End If
                            
                            '17/10/25 利子補給に伴う変更
                            wiFlgRishiHokyu = MBD010_借入利子補給金(w借入金.借入番号)
                            If wiFlgRishiHokyu <> 0 Then
                                wRs2("融資金額") = 0
                                wRs2("毎月返済額") = 0
                                wRs2("初回返済額") = 0
                                wRs2("最終返済額") = 0
                                wRs2("融資残高") = 0
                            End If
                            
                wRs2.Update
            End If
            
            wRs.MoveNext
        Loop
        wRs.Close
        Set wRs = Nothing
        
    wRs2.Close
    Set wRs2 = Nothing
    
    '----------------------------------------------------------------
    '                          ** 合計 計算 **
    '----------------------------------------------------------------
    
    '----------------------------------------------------------------
    '                       ** レコード　ソース **
    '----------------------------------------------------------------
    '** レコード　ソース **

    wstr = ""
    wstr = wstr & "Select "
    
    'wstr = wstr & " I.銀行番号 As GrpFld_Ginko,"
    'wstr = wstr & "G.銀行名 & ' 計' As G10_計名,"
    
    'セクションGR
    wstr = wstr & "K.銀行番号 As GrpFld_Ginko,"
    wstr = wstr & "G.金融機関番号 As GrpFld_Kinyu,"
    wstr = wstr & "B.部門番号 As GrpFld_Bumon,"
    wstr = wstr & " K.借入金種別区分 As GrpFld_KShubetu,"
    
    wstr = wstr & "G.銀行名 As I_銀行名,"
    wstr = wstr & "G.金融機関名 As I_金融機関名,"
    wstr = wstr & "B.部門名 As I_部門名,"
    wstr = wstr & "S.借入金種別名 As I_借入金種別名,"
    
    wstr = wstr & " I.借入番号 As I_借入番号,"
    wstr = wstr & " I.借入内容 As I_借入内容,"
    wstr = wstr & " I.入力借入計画番号 As I_借入計画番号,"
'    wstr = wstr & "'" & GRpt.借入 & "' As I_借入計画番号,"

    wstr = wstr & "K.sm区分 As I_SM区分,"
    wstr = wstr & " I.金融リストラ番号 As I_金融リストラ番号,"
    'wstr = wstr & " G.銀行名 As I_銀行名,"
    wstr = wstr & " IIF(I.有担保フラグ = " & P8.FCDbl(XMXA020_区分("有担フラグ", "有担保")) & ",'有担保','無担保') As I_担保,"
    
    wstr = wstr & " IIF(I.金利種別 = " & P8.FCDbl(XMXA020_区分("金利種別", "変動金利")) & ",'変動金利','固定金利') As I_金利種別,"
    wstr = wstr & " I.金利条件 As I_金利条件,"
    
    'wstr = wstr & " H.保証会社区分名 As I_保証会社区分名,"
    'wstr = wstr & " Y.融資区分名 As I_融資区分名,"
    
    wstr = wstr & " Format(I.入力残高年月,'" & Gfmt年月 & "') As I_残高年月,"
    wstr = wstr & " I.融資金額 As I_融資金額,"
    wstr = wstr & " I.融資残高 As I_融資残高,"
    wstr = wstr & " I.利率 As I_利率,"
    wstr = wstr & " IIF(I.変動利率フラグ=1,'*','') As I_利率フラグ,"
    wstr = wstr & " I.保証料率 As I_保証料率,"
    wstr = wstr & " I.返済単位月数 As I_返済単位月数,"
    wstr = wstr & " Format(I.実行日,'" & Gfmt年月日 & "') As I_実行日,"
    wstr = wstr & " Format(I.金融解約実行日,'" & Gfmt年月日 & "') As I_金融解約年月日,"
    wstr = wstr & " Format(I.解約実行日,'" & Gfmt年月日 & "') As I_解約年月日,"
    wstr = wstr & " Format(I.最終返済実行日,'" & Gfmt年月日 & "') As I_最終返済日,"
    wstr = wstr & " Format(I.初回返済年月,'" & Gfmt年月 & "') As I_初回返済年月,"
    wstr = wstr & " Format(I.最終返済年月,'" & Gfmt年月 & "') As I_最終返済年月,"
    
    wstr = wstr & " I.毎月返済額 As I_毎月返済額,"
    wstr = wstr & " I.初回返済額 As I_初回返済額,"
    wstr = wstr & " I.最終返済額 As I_最終返済額,"
    
    wstr = wstr & "IIF(手入力区分=1,'手入力','') As 手入力フラグ"
    
    wstr = wstr & " From (((DCDA030_借入一覧表 As I"
    wstr = wstr & " Inner Join " & wsTbl & " As K"
    wstr = wstr & " ON I.借入番号 = K.借入番号)"
    wstr = wstr & " Inner JOIN DAAA040_銀行マスタ As G"
    wstr = wstr & " ON K.銀行番号 = G.銀行番号)"
    'wstr = wstr & " LEFT JOIN DAAA100_保証会社区分 As H"
    'wstr = wstr & " ON K.保証会社区分 = H.保証会社区分)"
    'wstr = wstr & " LEFT JOIN DAAA110_融資区分 As Y"
    'wstr = wstr & " ON K.融資区分 = Y.融資区分"
    wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
    wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分)"
    wstr = wstr & " LEFT JOIN DAAA200_部門マスタ As B"
    wstr = wstr & " ON K.プロジェクト番号 = B.部門番号"

    'wOrder
    wOrder = "": FLG_Order = False
    For j = 1 To 2
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
        Call MX040_CsvOut_KARI
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
    Dim wism As Integer
'
    'シミュレーション内容
    wism = P8.FCDbl(Me.Detail.Controls("I_SM区分"))
    
    Me.Detail.Controls("I_SM区分") = ""
    
    If wism = 1 And P8.FCStr(Me.Detail.Controls("I_金融リストラ番号")) <> "" Then
        Me.Detail.Controls("I_SM区分") = "借入SM"
    ElseIf wism = 0 And P8.FCStr(Me.Detail.Controls("I_金融解約年月日")) <> "" Then
        Me.Detail.Controls("I_SM区分") = "解約SM"
    End If
'
    Call MXA030_ReportColor(Me.Detail.Controls("I_融資金額"))
    Call MXA030_ReportColor(Me.Detail.Controls("I_融資残高"))
    Call MXA030_ReportColor(Me.Detail.Controls("I_毎月返済額"))
    Call MXA030_ReportColor(Me.Detail.Controls("I_初回返済額"))
    Call MXA030_ReportColor(Me.Detail.Controls("I_最終返済額"))
    Call MXA030_ReportColor(Me.Detail.Controls("I_利率"))
    Call MXA030_ReportColor(Me.Detail.Controls("I_保証料率"))
    
    Me.Detail.Controls("I_融資金額") = Format(P8.FCDblRD(Me.Detail.Controls("I_融資金額") / w分母), "#,##0")
    Me.Detail.Controls("I_融資残高") = Format(P8.FCDblRD(Me.Detail.Controls("I_融資残高") / w分母), "#,##0")
    Me.Detail.Controls("I_毎月返済額") = Format(P8.FCDblRD(Me.Detail.Controls("I_毎月返済額") / w分母), "#,##0")
    Me.Detail.Controls("I_初回返済額") = Format(P8.FCDblRD(Me.Detail.Controls("I_初回返済額") / w分母), "#,##0")
    Me.Detail.Controls("I_最終返済額") = Format(P8.FCDblRD(Me.Detail.Controls("I_最終返済額") / w分母), "#,##0")
    Me.Detail.Controls("I_利率") = Format(P8.FCDblRD5(Me.Detail.Controls("I_利率")), "#,##0.00000")
    Me.Detail.Controls("I_保証料率") = Format(P8.FCDblRD5(Me.Detail.Controls("I_保証料率")), "#,##0.00000")
'
    '手入力の場合は表示しない
    If Me.Detail.Controls("手入力フラグ") <> "" Then
        'Me.Detail.Controls("I_利率") = ""
        'Me.Detail.Controls("I_利率フラグ") = ""
        'Me.Detail.Controls("I_保証料率") = ""
        Me.Detail.Controls("I_返済単位月数") = ""
        'Me.Detail.Controls("I_金利種別") = ""
        'Me.Detail.Controls("I_金利条件") = ""
    End If
'
End Sub

'------------------------------------------------
' GroupFooter1_BeforePrint
'------------------------------------------------
Private Sub GroupFooter1_BeforePrint()
'
    Call MXA030_ReportColor(Me.GroupFooter1.Controls("G10_融資金額"))
    Call MXA030_ReportColor(Me.GroupFooter1.Controls("G10_融資残高"))
    Call MXA030_ReportColor(Me.GroupFooter1.Controls("G10_毎月返済額"))

    Me.GroupFooter1.Controls("G10_融資金額") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_融資金額") / w分母), "#,##0")
    Me.GroupFooter1.Controls("G10_融資残高") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_融資残高") / w分母), "#,##0")
    Me.GroupFooter1.Controls("G10_毎月返済額") = Format(P8.FCDblRD(Me.GroupFooter1.Controls("G10_毎月返済額") / w分母), "#,##0")
'
End Sub

'------------------------------------------------
' GroupFooter2_BeforePrint
'------------------------------------------------
Private Sub GroupFooter2_BeforePrint()
'
    Call MXA030_ReportColor(Me.GroupFooter2.Controls("G20_融資金額"))
    Call MXA030_ReportColor(Me.GroupFooter2.Controls("G20_融資残高"))
    Call MXA030_ReportColor(Me.GroupFooter2.Controls("G20_毎月返済額"))

    Me.GroupFooter2.Controls("G20_融資金額") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_融資金額") / w分母), "#,##0")
    Me.GroupFooter2.Controls("G20_融資残高") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_融資残高") / w分母), "#,##0")
    Me.GroupFooter2.Controls("G20_毎月返済額") = Format(P8.FCDblRD(Me.GroupFooter2.Controls("G20_毎月返済額") / w分母), "#,##0")
'
End Sub

'------------------------------------------------
' ReportFooter_BeforePrint
'------------------------------------------------
Private Sub ReportFooter_BeforePrint()
'
    Call MXA030_ReportColor(Me.ReportFooter.Controls("G90_融資金額"))
    Call MXA030_ReportColor(Me.ReportFooter.Controls("G90_融資残高"))
    Call MXA030_ReportColor(Me.ReportFooter.Controls("G90_毎月返済額"))

    Me.ReportFooter.Controls("G90_融資金額") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_融資金額") / w分母), "#,##0")
    Me.ReportFooter.Controls("G90_融資残高") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_融資残高") / w分母), "#,##0")
    Me.ReportFooter.Controls("G90_毎月返済額") = Format(P8.FCDblRD(Me.ReportFooter.Controls("G90_毎月返済額") / w分母), "#,##0")
'
End Sub


