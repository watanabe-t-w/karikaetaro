Attribute VB_Name = "MBD020_借入金自動更新"
Option Explicit
'
Private Const pPROGRAM_ID As String = "MBD020_借入金自動更新"

'------------------------------------------------
' MBD020_借入金バッチ自動更新
'------------------------------------------------
Public Sub MBD020_借入金バッチ自動更新(p更新モード As String, _
                                      p推移表開始年度 As Integer, _
                                      p推移表区分 As String, _
                                      p売上計画番号 As String, _
                                      p借入計画番号 As String, _
                                      pリース計画番号 As String, _
                                      p設備計画番号 As String, _
                                      p金融リストラ As String, _
                                      p設備リストラ番号 As String, _
                                      pボトムアップ As Integer)         '07/02/12 V180
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    Dim wRs1 As ADODB.Recordset
    Dim wstr1 As String
    
    Dim w売上計画 As MAA910_売上計画
    
    Dim w整備売上計画番号(1500) As String
    Dim w整備cnt As Integer
    
    Dim w分類売上計画番号(1500) As String
    Dim w分類(1500) As Integer
    Dim w分類cnt As Integer
    
    Dim w取消フラグ2 As Integer
    
    Dim wcnt As Integer
    Dim w売上計画番号 As String
    
'    Dim J As Integer
    Dim wi年度 As Integer
    
    Dim ws基本 As String                       '5/6/29 V130
'
    On Error Resume Next
'
    Erase w整備売上計画番号()
    Erase w分類売上計画番号()
    Erase w分類()
'
    '--------------------------------------------
    '     マスタ整備処理ルーチン
    '--------------------------------------------
    
    '***資産売却に伴う処置
    G仮払消費税22F = 0                  ' 07/03/02 V180
    
    If p更新モード = "バッチ更新" Then
        '**取消売上計画番号セット**
        w整備cnt = 1
        wstr = ""
        wstr = wstr + "SELECT * FROM DBAA040_売上計画"
        wstr = wstr + " Where 取消フラグ = 1"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
            Do Until wRs.eof
            
                w売上計画番号 = wRs("売上計画番号")
                If wRs("売上計画番号") = Left$(w売上計画番号, 4) Then
                    w整備売上計画番号(w整備cnt) = w売上計画番号
                    w整備cnt = w整備cnt + 1
                End If
                                            
                wRs.MoveNext
            Loop
        wRs.Close
        Set wRs = Nothing
        
        w整備売上計画番号(w整備cnt) = "??????????"
        
        
        '**売上計画マスタ取消フラグ設定
        wstr = ""
        wstr = wstr + "SELECT * FROM DBAA040_売上計画"
        wstr = wstr + " Where 取消フラグ = 0"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
            Do Until wRs.eof
            
                w売上計画番号 = wRs("売上計画番号")
                For wcnt = 1 To w整備cnt
                    
                    If w整備売上計画番号(wcnt) = Left$(w売上計画番号, 4) Then
                        wRs("取消フラグ") = 1
                        Exit For
                    End If
                    
                Next
                
                wRs.Update
                wRs.MoveNext
            Loop
        wRs.Close
        Set wRs = Nothing
        
        '**設備計画マスタ取消フラグ設定
        wstr = ""
        wstr = wstr + "Select * From DBCA010_設備計画"
        wstr = wstr + " Where 取消フラグ = 0"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
            Do Until wRs.eof
            
                w売上計画番号 = wRs("設備計画番号")
                For wcnt = 1 To w整備cnt
                    
                    If w整備売上計画番号(wcnt) = w売上計画番号 Then
                        wRs("取消フラグ") = 1
                        Exit For
                    End If
                    
                Next
                
                wRs.Update
                wRs.MoveNext
            Loop
        wRs.Close
        Set wRs = Nothing
        
        
        '**売上計画取消テーブルＳＥＴ
        w整備cnt = 1
        wstr = ""
        wstr = wstr + "SELECT * FROM DBAA040_売上計画"
        wstr = wstr + " Where 取消フラグ = 1"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
            Do Until wRs.eof
            
                w整備売上計画番号(w整備cnt) = wRs("売上計画番号")
                w整備cnt = w整備cnt + 1
                                            
                wRs.MoveNext
            Loop
        wRs.Close
        Set wRs = Nothing
        
        w整備売上計画番号(w整備cnt) = "??????????"
        
        
        '**基本計画B取消フラグ設定
        wstr1 = ""
        wstr1 = wstr1 + "Select *"
        wstr1 = wstr1 + " From DBAA010_基本事業計画"
        wstr1 = wstr1 + " Where 取消フラグ = 1"
        Call AdoRecordsetOpen(GDb, wRs1, wstr1)
            Do Until wRs1.eof
            
                wi年度 = wRs1("事業計画開始年度")
                wstr = ""
                wstr = wstr + "Select *"
                wstr = wstr + " From DBAA010_基本事業計画B"
                wstr = wstr + " Where 事業計画開始年度 = " & wi年度
                Call AdoRecordsetOpen(GDb, wRs, wstr)
                    Do Until wRs.eof
                        wRs("取消フラグ") = 1
                        wRs.Update
                        wRs.MoveNext
                    Loop
                wRs.Close
                Set wRs = Nothing
                
                wRs1.MoveNext
            Loop
        wRs1.Close
        Set wRs1 = Nothing
         
        
        '**売上計画B取消フラグ設定
        wstr = ""
        wstr = wstr + "Select * From DBAA040_売上計画B"
        wstr = wstr + " Where 取消フラグ = 0"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
            Do Until wRs.eof
            
                w売上計画番号 = wRs("売上計画番号")
                For wcnt = 1 To w整備cnt
                    
                    If w整備売上計画番号(wcnt) = w売上計画番号 Then
                        wRs("取消フラグ") = 1
                        Exit For
                    End If
                    
                Next
                
                wRs.Update
                wRs.MoveNext
            Loop
        wRs.Close
        Set wRs = Nothing
   
        
        '**借入金取消フラグ設定
        wstr = ""
        wstr = wstr + "Select * From DBDA010_借入金"
        wstr = wstr + " Where 取消フラグ = 0"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
            Do Until wRs.eof
            
                w売上計画番号 = P8.FCStr(wRs("借入計画番号"))
                For wcnt = 1 To w整備cnt
                    
                    If w整備売上計画番号(wcnt) = w売上計画番号 Then
                        wRs("取消フラグ") = 1
                        Exit For
                    End If
                    
                Next
                
                wRs.Update
                wRs.MoveNext
            Loop
        wRs.Close
        Set wRs = Nothing
'
        Call MBD020_売上詳細卸業取消
        
        Call MBD020_売上詳細製造業取消
'
        '**銀行マスタ整理
        wstr = ""
        wstr = wstr + "Delete * From DAAA040_銀行マスタ"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**税率マスタ整理
        wstr = ""
        wstr = wstr + "Delete * From DAAA060_税率マスタ"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**基本事業計画整理
        wstr = ""
        wstr = wstr + "Delete * From DBAA010_基本事業計画"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**基本事業計画B整理
        wstr = ""
        wstr = wstr + "Delete * From DBAA010_基本事業計画B"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**事業登録整理
        wstr = ""
        wstr = wstr + "Delete * From DBAA020_事業登録"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**人員計画整理
        wstr = ""
        wstr = wstr + "Delete * From DBAA030_人員計画"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**売上計画整理
        wstr = ""
        wstr = wstr + "Delete * From DBAA040_売上計画"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**売上計画B整理
        wstr = ""
        wstr = wstr + "Delete * From DBAA040_売上計画B"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**売上実績整理
        If GTbl_売上実績 = "" Then
            GTbl_売上実績 = "DBBA010_売上実績"
        End If
        
        wstr = ""
'        wstr = wstr + "Delete * From DBBA010_売上実績"
        wstr = wstr + "Delete * From " & GTbl_売上実績
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**設備計画整理
        wstr = ""
        wstr = wstr + "Delete * From DBCA010_設備計画"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**固定資産勘定科目整理
        wstr = ""
        wstr = wstr + "Delete * From DAAC010_固定資産勘定科目マスタ"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**固定資産部門整理
        wstr = ""
        wstr = wstr + "Delete * From DAAC020_固定資産部門マスタ"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**借入金整理
        wstr = ""
        wstr = wstr + "Delete * From DBDA010_借入金"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**貸付金整理
        wstr = ""
        wstr = wstr + "Delete * From DBDA010_貸付金"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**プロジェクト整理
        wstr = ""
        wstr = wstr + "Delete * From DAAA080_プロジェクト"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**プロジェクト支店整理
        wstr = ""
        wstr = wstr + "Delete * From DAAA090_プロジェクト支店"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**保証会社区分
        wstr = ""
        wstr = wstr + "Delete * From DAAA100_保証会社区分"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
    
        '**融資区分
        wstr = ""
        wstr = wstr + "Delete * From DAAA110_融資区分"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**借入金明細TR整理
        wstr = ""
        wstr = wstr + "Delete * From DBDA010_借入金明細TR"
        wstr = wstr + " Where 取消フラグ = 1"
        wstr = wstr + " Or 取消フラグ２ = 1"
        GDb.Execute wstr
        
        '**借入金内入整理
        wstr = "DELETE U.*,K.借入番号"
        wstr = wstr + " FROM DBDA010_借入金内入1 As U"
        wstr = wstr + " LEFT JOIN DBDA010_借入金 As K"
        wstr = wstr + " ON U.借入番号 = K.借入番号"
        wstr = wstr + " WHERE K.借入番号 Is Null"
        GDb.Execute wstr
        
        wstr = "DELETE U.*,K.借入番号"
        wstr = wstr + " FROM DBDA010_借入金内入2 As U"
        wstr = wstr + " LEFT JOIN DBDA010_借入金 As K"
        wstr = wstr + " ON U.借入番号 = K.借入番号"
        wstr = wstr + " WHERE K.借入番号 Is Null"
        GDb.Execute wstr
        
        wstr = "DELETE U.*,K.借入番号"
        wstr = wstr + " FROM DBDA010_借入金内入3 As U"
        wstr = wstr + " LEFT JOIN DBDA010_借入金 As K"
        wstr = wstr + " ON U.借入番号 = K.借入番号"
        wstr = wstr + " WHERE K.借入番号 Is Null"
        GDb.Execute wstr
        
        wstr = "DELETE U.*,K.借入番号"
        wstr = wstr + " FROM DBDA010_借入金内入4 As U"
        wstr = wstr + " LEFT JOIN DBDA010_借入金 As K"
        wstr = wstr + " ON U.借入番号 = K.借入番号"
        wstr = wstr + " WHERE K.借入番号 Is Null"
        GDb.Execute wstr
        
        wstr = "DELETE U.*,K.借入番号"
        wstr = wstr + " FROM DBDA010_借入金内入5 As U"
        wstr = wstr + " LEFT JOIN DBDA010_借入金 As K"
        wstr = wstr + " ON U.借入番号 = K.借入番号"
        wstr = wstr + " WHERE K.借入番号 Is Null"
        GDb.Execute wstr
        
        wstr = "DELETE U.*,K.借入番号"
        wstr = wstr + " FROM DBDA010_借入金内入6 As U"
        wstr = wstr + " LEFT JOIN DBDA010_借入金 As K"
        wstr = wstr + " ON U.借入番号 = K.借入番号"
        wstr = wstr + " WHERE K.借入番号 Is Null"
        GDb.Execute wstr
        
        wstr = "DELETE U.*,K.借入番号"
        wstr = wstr + " FROM DBDA010_借入金内入7 As U"
        wstr = wstr + " LEFT JOIN DBDA010_借入金 As K"
        wstr = wstr + " ON U.借入番号 = K.借入番号"
        wstr = wstr + " WHERE K.借入番号 Is Null"
        GDb.Execute wstr
        
        '**貸付金明細TR整理
        wstr = ""
        wstr = wstr + "Delete * From DBDA010_貸付金明細TR"
        wstr = wstr + " Where 取消フラグ = 1"
        wstr = wstr + " Or 取消フラグ２ = 1"
        GDb.Execute wstr
        
        '**借入金借換整理
        wstr = "DELETE R.*,K.借入番号"
        wstr = wstr + " FROM DBDA010_借入金借換 As R"
        wstr = wstr + " LEFT JOIN DBDA010_借入金 As K"
        wstr = wstr + " ON R.借入番号 = K.借入番号"
        wstr = wstr + " WHERE K.借入番号 Is Null"
        GDb.Execute wstr
        
        '**リース整理
        wstr = ""
        wstr = wstr + "Delete * From DBDA010_リース"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**リース明細TR整理
        wstr = ""
        wstr = wstr + "Delete * From DBDA010_リース明細TR"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
         '**決算書科目マスタ整理
        wstr = ""
        wstr = wstr + "Delete * From DAAA050_決算書科目マスタ"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
         '**決算書科目TR整理
        wstr = ""
        wstr = wstr + "Delete * From DBAA050_決算書科目ＴＲ"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**売上詳細卸業整理
        wstr = ""
        wstr = wstr + "Delete * From DBAA060_売上詳細卸業"
        wstr = wstr + " Where (売上計画取消フラグ = 1) OR (分類取消フラグ = 1 )"
        GDb.Execute wstr
        
        '**売上詳細卸業整理B
        wstr = ""
        wstr = wstr + "Delete * From DBAA060_売上詳細卸業B"
        wstr = wstr + " Where (取消フラグ = 1) OR (取消フラグ2 = 1)"
        GDb.Execute wstr
        
        '**売上詳細製造業整理
        wstr = ""
        wstr = wstr + "Delete * From DBAA060_売上詳細製造業"
        wstr = wstr + " Where (売上計画取消フラグ = 1) OR (分類取消フラグ = 1 )"
        GDb.Execute wstr
        
        '**売上詳細製造業整理B
        wstr = ""
        wstr = wstr + "Delete * From DBAA060_売上詳細製造業B"
        wstr = wstr + " Where (取消フラグ = 1) OR (取消フラグ2 = 1)"
        GDb.Execute wstr
        
        '**本部経費振替整理  5/12/31 V130
        wstr = ""
        wstr = wstr + "Delete * From DBBA010_本部経費振替"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        '**基幹データ調整整理  5/12/31 V130
        wstr = ""
        wstr = wstr + "Delete * From DBBA010_基幹データ調整"
        wstr = wstr + " Where 取消フラグ = 1"
        GDb.Execute wstr
        
        
        '******売上詳細更新　テスト用
        'w売上計画番号 = "15-a"
        'Call MGA010_売上詳細卸業分類_合計作成(w売上計画番号)
        'Call MGA010_売上詳細更新_更新(w売上計画番号)
        
        'w売上計画番号 = "15-a"
        'Call MGA010_売上詳細製造業分類_合計作成(w売上計画番号)
        'Call MGA010_売上詳細更新_更新(w売上計画番号)
        
        '******損益分岐点分析　テスト用
        'p更新モード = "企業継続分岐点売上表"
        'p推移表開始年度 = 2003
        'p推移表区分 = "年次"
        'p売上計画番号 = "15-a"
        'p借入計画番号 = "15-a"
        'p設備計画番号 = "15-a"
        'p金融リストラ = ""
      
        
        'Call MBD040_損益分岐点更新P(p更新モード, _
        '                          p推移表開始年度, _
        '                          p推移表区分, _
        '                          p売上計画番号, _
        '                          p借入計画番号, _
        '                          p設備計画番号, _
        '                          p金融リストラ)

        
        
        '******基本事業計画より売上計画明細作成テスト用
        
        'wi年度 = 2003
        'wStr1 = ""
        'wStr1 = wStr1 + "Select *"
        'wStr1 = wStr1 + " From DBAA010_基本事業計画"
        'wStr1 = wStr1 + " Where 事業計画開始年度 = " & wi年度
        'Call AdoRecordsetOpen(GDb, wRs1, wStr1)
               
            'If Not wRs1.EOF Then
                'wStr2 = ""
                'wStr2 = wStr2 + "Select *"
                'wStr2 = wStr2 + " From DBAA040_売上計画"
                'wStr2 = wStr2 + " Where 取消フラグ = 0"
                'Call AdoRecordsetOpen(GDb, wRs2, wStr2)
                'Do Until wRs2.EOF
                    
                        
                        'wRs2("売上計画番号") = P8.FCStr(C年月日.年度開始年月日(wRs1("事業計画開始年度"), "e"))
                        'wRs2("売上計画内容") = " "
                        'wRs2("設備計画番号") = " "
                        'wRs2("売上計画開始年度") = wRs1("事業計画開始年度")
                        'wRs2("売上計画開始年月") = C年月日.年度開始年月日(Left(P8.FCStr(w売上計画番号), 2), "平成")
                        'wRs2("売上予算前年次") = wRs1("売上予算前年次")
                        'wRs2("人数前年次") = wRs1("人数前年次")
                       '
                        'For J = 1 To 10
                            'wRs2("売上予算" + CStr(J) + "年次") = wRs1("売上予算" + CStr(J) + "年次")
                            'wRs2("人数" + CStr(J) + "年次") = wRs1("人数" + CStr(J) + "年次")
                        'Next
                        'For J = 1 To 12
                            'wRs2("売上指数" + CStr(J) + "月度") = wRs1("売上指数" + CStr(J) + "月度")
                        'Next
                        
                        'wRs2("売上回収サイト") = wRs1("売上回収サイト")
                        'wRs2("売上回収1サイト") = wRs1("売上回収1サイト")
                        'wRs2("売上回収2サイト") = wRs1("売上回収2サイト")
                        'wRs2("売上回収3サイト") = wRs1("売上回収3サイト")
                        'wRs2("粗利率") = wRs1("粗利率")
                        'wRs2("粗利率1") = wRs1("粗利率1")
                        'wRs2("粗利率2") = wRs1("粗利率2")
                        'wRs2("粗利率3") = wRs1("粗利率3")
                        'wRs2("売上1構成比") = wRs1("売上1構成比")
                        'wRs2("売上2構成比") = wRs1("売上2構成比")
                        'wRs2("売上3構成比") = wRs1("売上3構成比")
                        'wRs2("売上達成率") = wRs1("売上達成率")
                        'wRs2("給与UP率") = wRs1("給与UP率")
                        'wRs2("賞与UP率") = wRs1("賞与UP率")
                        'wRs2("給与総額達成率") = wRs1("給与総額達成率")
                        'wRs2("賞与額達成率") = wRs1("賞与額達成率")
                        'wRs2("固定経費達成率") = wRs1("固定経費達成率")
                        'wRs2("変動経費1達成率") = wRs1("変動経費1達成率")
                        'wRs2("変動経費2達成率") = wRs1("変動経費2達成率")
                        'wRs2("変動経費3達成率") = wRs1("変動経費3達成率")
                        'wRs2("その他経費1達成率") = wRs1("その他経費1達成率")
                        'wRs2("保険積立達成率") = wRs1("保険積立達成率")
                        'wRs2("営業外収益達成率") = wRs1("営業外収益達成率")
                        'wRs2("営業外費用達成率") = wRs1("営業外費用達成率")
                        'wRs2("減価償却費達成率") = wRs1("減価償却費達成率")
                        'wRs2("支払利息達成率") = wRs1("支払利息達成率")
                        'wRs2("給与総額") = wRs1("給与総額")
                        'wRs2("新人給与月額") = wRs1("新人給与月額")
                        'wRs2("賞与額") = wRs1("賞与額")
                        'wRs2("新人賞与額") = wRs1("新人賞与額")
                        'wRs2("固定経費") = wRs1("固定経費")
                        'wRs2("変動経費1") = wRs1("変動経費1")
                        'wRs2("変動経費2") = wRs1("変動経費2")
                        'wRs2("変動経費3") = wRs1("変動経費3")
                        'wRs2("その他経費1") = wRs1("その他経費1")
                        'wRs2("保険積立") = wRs1("保険積立")
                        'wRs2("営業外収益") = wRs1("営業外収益")
                        'wRs2("営業外費用") = wRs1("営業外費用")
                        'wRs2("減価償却費") = wRs1("減価償却費")
                        'wRs2("支払利息") = wRs1("支払利息")
                        'wRs2("手持資金") = wRs1("手持資金")
                        'wRs2("その他資金1") = wRs1("その他資金1")
                        'wRs2("その他資金2") = wRs1("その他資金2")
                        'wRs2("取消フラグ") = wRs1("取消フラグ")
                   '
                        'w売上計画 = MBB010_売上計画データセット(wRs2)
                       '
                        
                
                        'Call MBB010_売上計画テーブル作成(w売上計画)
  
                        
                        'wStr = ""
                        'wStr = wStr + "Delete * From DCAA010_売上計画明細"
                        'GDb.Execute wStr
                        'Call MBB010_売上計画明細作成
                        'Exit Do
                        
                'Loop
                 
                
                    'wRs2.Close
                    
            'End If
            
         
        'wRs1.Close
        
        '******V103A TO V110 DATA CVT テスト用
        
        'Call ZCV000_103A_110変換
        
        Exit Sub
        
    End If
'
    On Error GoTo 0
'

'
    '--------------------------------------------
    '     メイン処理ルーチン
    '--------------------------------------------
'
    On Error GoTo MBD020_借入金バッチ自動更新_ERR
'
    If p更新モード = "バッチ更新" Then
        wstr = "SELECT * FROM DBAA040_売上計画"
        wstr = wstr + " Where 取消フラグ = 0"
        Call AdoRecordsetOpen(GDb, wRs, wstr)
            Do Until wRs.eof
            
                w売上計画 = MBB010_売上計画データセット(wRs)
   
                Call MBD020_借入ファイル更新(p更新モード, p推移表区分, w売上計画, _
                                             w売上計画.売上計画番号, p金融リストラ, _
                                             w売上計画.売上計画番号, _
                                             w売上計画.設備計画番号, p設備リストラ番号, p推移表開始年度) ' 07/02/12 V180
                wRs.MoveNext
            Loop
        wRs.Close
        Set wRs = Nothing
    
    ElseIf p更新モード = "自動更新" _
            Or p更新モード = "損益資金計画推移表" Or p更新モード = "損益計画推移表" _
            Or p更新モード = "損益予実対比表" Or p更新モード = "損益予実対比表予算" _
            Or p更新モード = "経営計画支援表" Or p更新モード = "経営計画支援表２" _
            Or p更新モード = "企業継続分岐点売上表" Or p更新モード = "損益分岐点売上表" Then
            
        ws基本 = Left$(p売上計画番号, 2)                   '5/6/29 V130
        If ws基本 <> p売上計画番号 Then                    '5/6/29 V130
            wstr = "SELECT * FROM DBAA040_売上計画"
            wstr = wstr + " WHERE  売上計画番号= '" & p売上計画番号 & "'"
            Call AdoRecordsetOpen(GDb, wRs, wstr)
        
                w売上計画 = MBB010_売上計画データセット(wRs)
            wRs.Close
            Set wRs = Nothing
        Else
            wi年度 = P8.FCDbl(C年月日.年度開始年月日(ws基本, "y"))    '5/6/29 V130
        
            wstr = wstr + "SELECT * FROM DBAA010_基本事業計画"      '5/6/29 V130
            wstr = wstr + " WHERE  事業計画開始年度= " & wi年度      '5/6/29 V130
            Call AdoRecordsetOpen(GDb, wRs, wstr)                   '5/6/29 V130
        
                w売上計画 = MBB010_基本計画データセット(wRs)
            wRs.Close
            Set wRs = Nothing
        End If
        
        Call MBD020_借入ファイル更新(p更新モード, p推移表区分, w売上計画, _
                                     p借入計画番号, p金融リストラ, _
                                     pリース計画番号, _
                                     p設備計画番号, p設備リストラ番号, pボトムアップ, p推移表開始年度) '07/02/12 V180
    End If
'
    Erase w整備売上計画番号()
    Erase w分類売上計画番号()
    Erase w分類()
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD020_借入金バッチ自動更新_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD020_借入金バッチ自動更新() でエラー" + vbCrLf + vbCrLf + _
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
' MBD020_借入ファイル更新
'------------------------------------------------
Private Sub MBD020_借入ファイル更新(p更新モード As String, _
                            p推移表区分 As String, _
                            p売上計画 As MAA910_売上計画, _
                            p借入計画番号 As String, _
                            p金融リストラ As String, _
                            pリース計画番号 As String, _
                            p設備計画番号 As String, _
                            p設備リストラ番号 As String, _
                            pボトムアップ As Integer, _
                            Optional p推移表開始年度 As Integer) '06/05/19 V160
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    Dim wRs1 As ADODB.Recordset
    Dim wstr1 As String
    Dim j As Integer
    Dim k As Integer
   
    Dim w設備計画 As MAA910_設備計画
    Dim wRet設備計画作成 As MBC010_設備計画テーブル作成リターン
    
    Dim w借入金 As MAA910_借入金
    Dim wリース As MAA910_リース    '08/04/15 V182
    Dim w解約保証料戻 As Double 'MBC010_借入金テーブル作成リターン
    
    Dim wRet資金調達 As MBA010_資金調達リターン
   
    Dim w出力用借入金テーブル() As MAA910_借入金
    Dim w配列数 As Integer
   
    Dim w基本事業計画 As MAA910_基本事業計画
    
    Dim wベンチャ As String
    Dim wベンチャcode As String
    
    Dim w借入番号 As String
    Dim w自動借入連番 As Long
    
    Dim w設備計画番号 As String
    Dim ww設備計画番号 As String     '5/6/29 V130
    Dim www設備計画番号 As String    '5/8/30 V129
    Dim wFILE名 As String
    Dim wFILE名2 As String  '2004/10/6 V120
    Dim w対象年次 As Integer
    Dim w対象最終年次 As Integer
    
    Dim ws基本 As String                       '5/6/29 V130
    
    Dim w期首借入年度 As String                 '5/10/9 V129 支店への貸付　リストラ番号
    Dim w支店貸付 As String                     '5/10/9 V129 支店への貸付　リストラ番号
    Dim w全社借入 As String                     '5/10/9 V129 支店への借入　リストラ番号
    
    Dim w借入貸付sw As Integer                  '06/04/02 V150
'
    On Error GoTo MBD020_借入ファイル更新_ERR
'
    wベンチャ = Left$(p売上計画.売上計画番号, 4)
    wベンチャcode = StrConv(Right$(wベンチャ, 1), 2)
    
    ws基本 = Left$(p売上計画.売上計画番号, 2)                   '5/6/29 V130
    If ws基本 = p売上計画.売上計画番号 Then                     '5/6/29 V130
        wベンチャcode = "a"                                     '5/7/5  V130
    End If                                                      '5/7/5  V130
    
    w対象年次 = 0    '***10年間対象
    If G会議資金調達 = "連結" Then                            '06/10/18 V170
        Select Case p推移表区分
            Case "月次":    w対象最終年次 = 1 + 2
            Case "四半期":  w対象最終年次 = 3 + 2
            Case "半期":    w対象最終年次 = 5 + 2
            Case "年次":    w対象最終年次 = 10                    '06/09/12 V170
        End Select
    Else
        Select Case p推移表区分
            Case "月次":    w対象最終年次 = 1
            Case "四半期":  w対象最終年次 = 3
            Case "半期":    w対象最終年次 = 5
            Case "年次":    w対象最終年次 = 10                    '06/09/12 V170
        End Select
    End If
    
    w対象最終年次 = w対象最終年次 + p推移表開始年度 - p売上計画.売上計画開始年度
    
    ' -----------------------------------------
    '              科目テーブル クリア
    ' -----------------------------------------
    If pボトムアップ = 0 Then           '06/05/19 V160
        Call MAA500_科目集計クリア
    End If                              '06/05/19 V160
        
    ' -----------------------------------------
    '              基本事業計画 セット
    ' -----------------------------------------
    wstr = ""
    wstr = wstr + "SELECT * FROM DBAA010_基本事業計画"
    wstr = wstr + " WHERE  事業計画開始年度= " & p売上計画.売上計画開始年度
    Call AdoRecordsetOpen(GDb, wRs, wstr)
    
        w基本事業計画.事業計画開始年度 = wRs("事業計画開始年度")
        w基本事業計画.売上予算X年次(0) = wRs("売上予算前年次")
        w基本事業計画.人数X年次(0) = wRs("人数前年次")
        
        For j = 1 To 10
            w基本事業計画.売上予算X年次(j) = wRs("売上予算" + CStr(j) + "年次")
            w基本事業計画.人数X年次(j) = wRs("人数" + CStr(j) + "年次")
            If j = 10 Then              '06/09/12 V170
                w基本事業計画.売上予算X年次(11) = wRs("売上予算" + CStr(j) + "年次") '06/09/12 V170
                w基本事業計画.人数X年次(11) = wRs("人数" + CStr(j) + "年次")         '06/09/12 V170
            End If                      '06/09/12 V170
        Next
        For j = 1 To 12
            w基本事業計画.売上指数X月度(j) = wRs("売上指数" + CStr(j) + "月度")
        Next
        
        w基本事業計画.粗利率 = wRs("粗利率")
        w基本事業計画.売上達成率 = wRs("売上達成率")
        w基本事業計画.売上1達成率 = wRs("売上1達成率")
        w基本事業計画.売上2達成率 = wRs("売上2達成率")
        w基本事業計画.売上3達成率 = wRs("売上3達成率")
        w基本事業計画.給与up率 = wRs("給与Up率")
        w基本事業計画.賞与up率 = wRs("賞与Up率")
        w基本事業計画.給与総額 = wRs("給与総額")
        w基本事業計画.新人給与月額 = wRs("新人給与月額")
        w基本事業計画.賞与額 = wRs("賞与額")
        w基本事業計画.新人賞与額 = wRs("新人賞与額")
        w基本事業計画.固定経費 = wRs("固定経費")
        w基本事業計画.変動経費1 = wRs("変動経費1")
        w基本事業計画.変動経費2 = wRs("変動経費2")
        w基本事業計画.変動経費3 = wRs("変動経費3")
        w基本事業計画.その他経費1 = wRs("その他経費1")
        w基本事業計画.定期積金 = wRs("定期積金")                '06/03/11 V150
        w基本事業計画.協力積立金 = wRs("協力積立金")            '06/03/11 V150
        w基本事業計画.保険積立 = wRs("保険積立")
        w基本事業計画.受取リベート = wRs("受取リベート")        '06/03/11 V150
        w基本事業計画.支払リベート = wRs("支払リベート")        '06/03/11 V150
        w基本事業計画.営業外収益 = wRs("営業外収益")
        w基本事業計画.営業外費用 = wRs("営業外費用")
        w基本事業計画.減価償却費 = wRs("減価償却費")
        w基本事業計画.支払利息 = wRs("支払利息")
        w基本事業計画.手持資金 = wRs("手持資金")
        w基本事業計画.定期積金残 = wRs("定期積金残")            '06/03/11 V150
        w基本事業計画.協力積立金残 = wRs("協力積立金残")        '06/03/11 V150
        w基本事業計画.その他資金1 = wRs("その他資金1")
        w基本事業計画.その他資金2 = wRs("その他資金2")
        w基本事業計画.その他資金3 = wRs("その他資金3")
        w基本事業計画.売掛残高 = wRs("売掛残高")
        w基本事業計画.買掛残高 = wRs("買掛残高")
        w基本事業計画.投資債権残 = wRs("投資債権残")            '06/03/11 V150
        w基本事業計画.投資債務残 = wRs("投資債務残")            '06/03/11 V150
        w基本事業計画.未収入金残 = wRs("未収入金残")            '06/03/11 V150
        w基本事業計画.未払費用残 = wRs("未払費用残")            '06/03/11 V150
        w基本事業計画.期末在庫 = wRs("期末在庫")      '5/5/5 V127
        w基本事業計画.前期繰越利益 = wRs("前期繰越利益")          '5/9/16 V129
        w基本事業計画.その他債権残高 = wRs("その他債権残高")      '5/5/5 V127
        w基本事業計画.その他債権残高2 = wRs("その他債権残高2")    '06/03/11 V150
        w基本事業計画.その他債務残高 = wRs("その他債務残高")      '5/5/5 V127
        w基本事業計画.その他債務残高2 = wRs("その他債務残高2")    '06/03/11 V150
        
        w基本事業計画.取消フラグ = wRs("取消フラグ")
    
    wRs.Close
    Set wRs = Nothing
    
    ' -----------------------------------------
    '              借入金ファイル クリア
    ' -----------------------------------------
    If G基本情報.企業区分 <> "本部" And G基本情報.企業区分 <> "連結本部" Then          '06/09/262 V170
    
        If p更新モード = "自動更新" Or p更新モード = "バッチ更新" Or _
               (p更新モード <> "損益予実対比表予算" And p借入計画番号 <> "") Then
           wstr = ""
           wstr = wstr + "Delete * From DBDA010_借入金"
           wstr = wstr + " Where 借入計画番号 ='" & p売上計画.売上計画番号 & "'"
           wstr = wstr + " And Sm区分 = 1"
           GDb.Execute wstr
        End If
    
    End If                                                  '06/05/02 V150
    
        
    ' -----------------------------------------
    '      売上計画マスタより 科目テーブル作成
    ' -----------------------------------------
    '** 売上計画テーブル セット **
        Call MBB010_売上計画テーブル作成(p売上計画)     '06/08/09 V160
    If pボトムアップ = 0 Then               '08/05/19 V160
        '** 売上計画テーブル セット **
        'Call MBB010_売上計画テーブル作成(p売上計画)    '06/08/09 V160
    
        '** 科目テーブル セット **
        Call MBA010_科目集計セット_売上計画(p更新モード, p売上計画)
    End If                                  '06/05/19 V160

    ' -----------------------------------------
    '      設備計画マスタより 科目テーブル作成
    ' -----------------------------------------
  If pボトムアップ = 0 Then               '06/05/19 V160

'    wStr = ""
'    wStr = wStr + "Select * From DBCA010_設備計画"
'    wStr = wStr + " Where Sm区分 = 0"
'    wStr = wStr + "      Or ( 設備計画番号 = '" & p設備計画番号 & "'"
'    wStr = wStr + "      And Format(設備年月,'yyyymmdd') > '" & Format(Gコントロール.最終実績年月, "yyyymmdd") & "')"
'
    w設備計画番号 = Left$(p設備計画番号, 6)
    ww設備計画番号 = Left$(p設備計画番号, 4)    '5/6/29 V130
    www設備計画番号 = Left$(p設備計画番号, 2)   '5/8/30 V129
    
    wstr = ""
    wstr = wstr + "Select * From DBCA010_設備計画"
    If p更新モード = "損益予実対比表予算" Then
        wstr = wstr + " Where (設備計画番号 = '" & p設備計画番号 & "'"
        wstr = wstr + " Or 設備計画番号 = '" & w設備計画番号 & "'"      '5/6/29 V130
        wstr = wstr + " Or 設備計画番号 = '" & ww設備計画番号 & "'"     '5/6/29 V130
        wstr = wstr + " Or 設備計画番号 = '" & www設備計画番号 & "')"   '5/8/30 V129
        wstr = wstr + " And 取消フラグ = 0"
    Else
        If wベンチャcode < "a" Or wベンチャcode > "z" Then
            wstr = wstr + " Where "
        Else
            wstr = wstr + " Where (Sm区分 = 0 And 取消フラグ = 0) Or "
        End If
        wstr = wstr + "       ((( 設備計画番号 = '" & p設備計画番号 & "'"
        wstr = wstr + " Or 設備計画番号 = '" & w設備計画番号 & "'"     '5/6/29 V130
        wstr = wstr + " Or 設備計画番号 = '" & ww設備計画番号 & "'"    '5/6/29 V130
        wstr = wstr + " Or 設備計画番号 = '" & www設備計画番号 & "')"  '5/8/30 V129
        If wベンチャcode < "a" Or wベンチャcode > "z" Then
            wstr = wstr + ")"
        Else
            wstr = wstr + "      And Format(設備年月,'yyyymmdd') > '" & Format(Gコントロール.最終実績年月, "yyyymmdd") & "')"
        End If
        wstr = wstr + "      And 取消フラグ = 0)"
    End If
    
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            w設備計画 = MBC010_設備計画データセット(wRs)
     
            '** 設備計画テーブル セット **
            wRet設備計画作成 = MBC010_設備計画テーブル作成(w設備計画, p設備リストラ番号) ' 07/02/12 V180
     
            '** 科目テーブル セット **
            Call MBA010_科目集計セット_設備計画(p更新モード, w設備計画, p設備リストラ番号, p売上計画) ' 07/02/12 V180
     
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
     
  End If                           '06/05/19 V160
   
    ' -----------------------------------------
    '       借入金マスタより 科目テーブル作成
    ' -----------------------------------------
    w借入貸付sw = 1                                         '06/04/02 V150
    '**ベンチャーの時で借入金自動登録の場合　対象外
    w期首借入年度 = Left$(p売上計画.売上計画番号, 2)        '5/10/9 V129
    w支店貸付 = "支店貸付"                                  '5/10/9 V129
    w全社借入 = "全社借入"                                  '5/10/17 V129
    
    If (wベンチャcode < "a" Or wベンチャcode > "z") And _
        (p更新モード = "自動更新" Or p更新モード = "バッチ更新") Then
    Else
    
STA借入貸付:                                                '06/04/02 V150
        wstr = ""
        If w借入貸付sw = 1 Then                             '06/04/02 V150
          If ((G基本情報.企業区分 = "本部" And pボトムアップ = 1) _
              And (p更新モード = "企業継続分岐点売上表" Or p更新モード = "損益分岐点売上表")) _
          Or ((G基本情報.企業区分 = "連結本部" And pボトムアップ = 2) _
              And (p更新モード = "企業継続分岐点売上表" Or p更新モード = "損益分岐点売上表")) Then '06/09/29 V170
            wstr = wstr + "Select * From DBDA010_分岐点借入金"  '06/07/31 V160
          Else                                                  '06/07/31 V160
            wstr = wstr + "Select * From DBDA010_借入金"
          End If                                                '06/07/31 V160
        
        Else                                                '06/04/02 V150
          If ((G基本情報.企業区分 = "本部" And pボトムアップ = 1) _
              And (p更新モード = "企業継続分岐点売上表" Or p更新モード = "損益分岐点売上表")) _
          Or ((G基本情報.企業区分 = "連結本部" And pボトムアップ = 2) _
              And (p更新モード = "企業継続分岐点売上表" Or p更新モード = "損益分岐点売上表")) Then '06/09/29 V170
            wstr = wstr + "Select * From DBDA010_分岐点貸付金"  '06/07/31 V160
          Else                                                  '06/07/31 V160
            wstr = wstr + "Select * From DBDA010_貸付金"    '06/04/02 V150
          End If                                                '06/07/31 V160
        End If                                              '06/04/02 V150
        
        wstr = wstr + " Where 手入力区分 <> 2"              ' 07/02/12 V180
        
        '**ベンチャーの場合
        If wベンチャcode < "a" Or wベンチャcode > "z" Then
            p金融リストラ = ""   '**ﾍﾞﾝﾁﾔｰのとき　金融リストラは　無し　2003/6/3
            
            wstr = wstr + " And ("                          ' 07/02/12 V180
            If p金融リストラ <> "" Then
                wstr = wstr + " (金融リストラ番号 = '" & p金融リストラ & "'"
                wstr = wstr + " And 取消フラグ = 0) Or  "
            End If
            
            If p更新モード = "損益資金計画推移表" Or p更新モード = "損益計画推移表" _
                Or p更新モード = "損益予実対比表" Or p更新モード = "損益予実対比表予算" _
                Or p更新モード = "経営計画支援表" Or p更新モード = "経営計画支援表２" _
                Or p更新モード = "企業継続分岐点売上表" Or p更新モード = "損益分岐点売上表" Then
                
                wstr = wstr + " (借入計画番号 = '" & p借入計画番号 & "'"
                wstr = wstr + " And 借入計画番号 <> ''"
                
                'If G支店 <> 1 Then '支店の時　資金自動調達　有効　5/7/28 V128 西川産業向け
                    'wstr = wstr + "      And Format(実行日,'yyyymmdd') > '" & Format(Gコントロール.最終実績年月, "yyyymmdd") & "'"
                'End If
                
                If G基本情報.企業区分 = "連結本部" Then         '06/09/22 V170
                    wstr = wstr + " And Format(実行日,'yyyymmdd') > '" & Format(Gコントロール.最終実績年月, "yyyymmdd") & "'"
                End If                                    '17/8/17 V128
                
                wstr = wstr + " And 取消フラグ = 0)"
            End If
        Else
        '**会社全体の場合
            wstr = wstr + " And ((Sm区分 = 0 And 取消フラグ = 0) "  ' 07/02/12 V180
            
            If G基本情報.企業区分 = "本部" Or G基本情報.企業区分 = "連結本部" Then      '06/09/29 V170                          '06/05/02 V150
'                wstr = wstr + " Or (借入計画番号 = '" & p借入計画番号 & "'"  '06/05/02 V150
'                wstr = wstr + " And 取消フラグ = 0)"                         '06/05/02 V150
            
                wstr = wstr + " Or (借入計画番号='' And sm区分 = 0 And 取消フラグ = 0)"
                wstr = wstr + " Or (借入計画番号 = '" & p借入計画番号 & "' And 借入計画番号 <> ''"
                wstr = wstr + " And sm区分 = 0 And 取消フラグ = 0)"
            
            End If                                                          '06/05/02 V150
            
            If p金融リストラ <> "" Then
                wstr = wstr + " Or (金融リストラ番号 = '" & p金融リストラ & "'"
                wstr = wstr + " And 取消フラグ = 0)"
            End If
            
                wstr = wstr + " Or (金融リストラ番号 = '" & w期首借入年度 & "'"    '5/10/9 V129
                wstr = wstr + " And 取消フラグ = 0)"                             '5/10/9 V129
                
                wstr = wstr + " Or (金融リストラ番号 = '" & w支店貸付 & "'"        '5/10/9 V129
                wstr = wstr + " And 取消フラグ = 0)"                             '5/10/9 V129
                
                wstr = wstr + " Or (金融リストラ番号 = '" & w全社借入 & "'"        '5/10/9 V129
                wstr = wstr + " And 取消フラグ = 0)"                             '5/10/9 V129
            
            If p更新モード = "損益資金計画推移表" Or p更新モード = "損益計画推移表" _
                Or p更新モード = "損益予実対比表" Or p更新モード = "損益予実対比表予算" _
                Or p更新モード = "経営計画支援表" Or p更新モード = "経営計画支援表２" _
                Or p更新モード = "企業継続分岐点売上表" Or p更新モード = "損益分岐点売上表" Then
                
                wstr = wstr + "    Or  (借入計画番号 = '" & p借入計画番号 & "'"
                wstr = wstr + "    And 借入計画番号 <> ''"
                
                'If G支店 <> 1 Then '支店の時　資金自動調達　有効　5/7/28 V128 西川産業向け
                    'wstr = wstr + "      And Format(実行日,'yyyymmdd') > '" & Format(Gコントロール.最終実績年月, "yyyymmdd") & "'"
                'End If
                
                If G基本情報.企業区分 = "連結本部" Then         '06/09/22 V170
                    wstr = wstr + " And Format(実行日,'yyyymmdd') > '" & Format(Gコントロール.最終実績年月, "yyyymmdd") & "'"
                End If                                    '17/8/17 V128
                
                wstr = wstr + " And 取消フラグ = 0)"
            End If
        End If
        
        wstr = wstr + ")"                                       ' 07/02/12 V180
        
        
        ''06/10/02 V170
        'If (G基本情報.企業区分 = "連結本部" And pボトムアップ = 2) _
        'And (p更新モード = "企業継続分岐点売上表" Or p更新モード = "損益分岐点売上表") Then
        '    If w借入貸付sw = 1 Then
        '        wstr = wstr + " UNION Select * From DBDA010_借入金"
        '    Else
        '        wstr = wstr + " UNION Select * From DBDA010_貸付金"
        '    End If
        '        wstr = wstr + " Where 手入力区分 <> 2"          ' 07/02/12 V180
        '        wstr = wstr + " And ((借入計画番号='' And sm区分 = 0 And 取消フラグ = 0)"   ' 07/02/12 V180
        '        wstr = wstr + " Or (借入計画番号 = '" & p借入計画番号 & "' And 借入計画番号 <> ''"
        '        wstr = wstr + " And sm区分 = 0 And 取消フラグ = 0))"    ' 07/02/12 V180
        'End If

借入金2テーブル:        '09/04/30 V189 上記union分の処理
        Call AdoRecordsetOpen(GDb, wRs, wstr)
            Do Until wRs.eof
                w借入金 = MBD010_借入データセット(wRs)
         
                '** 借入金テーブル セット **
                If w借入金.手入力区分 <> P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then                  ' 07/02/12 V180
                    Call MBA010_科目集計セット_借入金手入力(p更新モード, w借入金, p売上計画) ' 07/02/12 V180
                Else                                            ' 07/02/12 V180
                    Call MBD010_借入金テーブル作成(p金融リストラ, w借入金)

                    '** 科目テーブル セット **
                    Call MBA010_科目集計セット_借入金(p更新モード, w借入金, p売上計画, p金融リストラ)
                End If                                          ' 07/02/12 V180
                
                wRs.MoveNext
            Loop
        wRs.Close
        Set wRs = Nothing
        
        
        '09/04/30 V189
        wstr = ""
        If (G基本情報.企業区分 = "連結本部" And pボトムアップ = 2) _
        And (p更新モード = "企業継続分岐点売上表" Or p更新モード = "損益分岐点売上表") Then
            If w借入貸付sw = 1 Then
                wstr = wstr + "Select * From DBDA010_借入金"
            Else
                wstr = wstr + "Select * From DBDA010_貸付金"
            End If
                wstr = wstr + " Where 手入力区分 <> 2"          ' 07/02/12 V180
                wstr = wstr + " And ((借入計画番号='' And sm区分 = 0 And 取消フラグ = 0)"   ' 07/02/12 V180
                wstr = wstr + " Or (借入計画番号 = '" & p借入計画番号 & "' And 借入計画番号 <> ''"
                wstr = wstr + " And sm区分 = 0 And 取消フラグ = 0))"    ' 07/02/12 V180
        End If
        
        '09/04/30 V189
        If wstr <> "" Then
            GoTo 借入金2テーブル
        End If
        
        If w借入貸付sw = 1 Then                         '06/04/02 V150
            w借入貸付sw = 2                             '06/04/02 V150
            
            GoTo STA借入貸付                            '06/04/02 V150
        End If                                          '06/04/02 V150
        
    End If
    
    
    
    
    ' -----------------------------------------
    '       リースマスタより 科目テーブル作成
    ' -----------------------------------------
    w借入貸付sw = 1                                         '06/04/02 V150
    '**ベンチャーの時で借入金自動登録の場合　対象外
    w期首借入年度 = Left$(p売上計画.売上計画番号, 2)        '5/10/9 V129
    w支店貸付 = "支店貸付"                                  '5/10/9 V129
    w全社借入 = "全社借入"                                  '5/10/17 V129
    
    


    If (wベンチャcode < "a" Or wベンチャcode > "z") And _
        (p更新モード = "自動更新" Or p更新モード = "バッチ更新") Then
    Else
    
                                                '06/04/02 V150
        wstr = ""
        wstr = wstr + "Select * From DBDA010_リース"
         
        wstr = wstr + " Where 手入力区分 <> 2"              ' 07/02/12 V180
        
        '**ベンチャーの場合
        If wベンチャcode < "a" Or wベンチャcode > "z" Then
               
            If p更新モード = "損益資金計画推移表" Or p更新モード = "損益計画推移表" _
                Or p更新モード = "損益予実対比表" Or p更新モード = "損益予実対比表予算" _
                Or p更新モード = "経営計画支援表" Or p更新モード = "経営計画支援表２" _
                Or p更新モード = "企業継続分岐点売上表" Or p更新モード = "損益分岐点売上表" Then
                
                wstr = wstr + " (リース計画番号 = '" & pリース計画番号 & "'"
                wstr = wstr + " And リース計画番号 <> ''"
                
                  
                If G基本情報.企業区分 = "連結本部" Then         '06/09/22 V170
                    wstr = wstr + " And Format(実行日,'yyyymmdd') > '" & Format(Gコントロール.最終実績年月, "yyyymmdd") & "'"
                End If                                    '17/8/17 V128
                
                wstr = wstr + " And 取消フラグ = 0)"
            End If
        Else
        '**会社全体の場合
            wstr = wstr + " And ((Sm区分 = 0 And 取消フラグ = 0) "  ' 07/02/12 V180
            
            If G基本情報.企業区分 = "本部" Or G基本情報.企業区分 = "連結本部" Then      '06/09/29 V170                          '06/05/02 V150
                wstr = wstr + " Or (リース計画番号='' And sm区分 = 0 And 取消フラグ = 0)"
                wstr = wstr + " Or (リース計画番号 = '" & pリース計画番号 & "' And リース計画番号 <> ''"
                wstr = wstr + " And sm区分 = 0 And 取消フラグ = 0)"
            
            End If                                                          '06/05/02 V150
            
                  
            If p更新モード = "損益資金計画推移表" Or p更新モード = "損益計画推移表" _
                Or p更新モード = "損益予実対比表" Or p更新モード = "損益予実対比表予算" _
                Or p更新モード = "経営計画支援表" Or p更新モード = "経営計画支援表２" _
                Or p更新モード = "企業継続分岐点売上表" Or p更新モード = "損益分岐点売上表" Then
                
                wstr = wstr + "    Or  (リース計画番号 = '" & pリース計画番号 & "'"
                wstr = wstr + "    And リース計画番号 <> ''"
                
                
                If G基本情報.企業区分 = "連結本部" Then         '06/09/22 V170
                    wstr = wstr + " And Format(実行日,'yyyymmdd') > '" & Format(Gコントロール.最終実績年月, "yyyymmdd") & "'"
                End If                                    '17/8/17 V128
                
                wstr = wstr + " And 取消フラグ = 0)"
            End If
        End If
        
        wstr = wstr + ")"                                       ' 07/02/12 V180
        
        '06/10/02 V170
        If (G基本情報.企業区分 = "連結本部" And pボトムアップ = 2) _
        And (p更新モード = "企業継続分岐点売上表" Or p更新モード = "損益分岐点売上表") Then
            
            wstr = wstr + " UNION Select * From DBDA010_リース"
            wstr = wstr + " Where 手入力区分 <> 2"          ' 07/02/12 V180
            wstr = wstr + " And ((リース計画番号='' And sm区分 = 0 And 取消フラグ = 0)"   ' 07/02/12 V180
            wstr = wstr + " Or (リース計画番号 = '" & pリース計画番号 & "' And リース計画番号 <> ''"
            wstr = wstr + " And sm区分 = 0 And 取消フラグ = 0))"    ' 07/02/12 V180
        End If
        
        Call AdoRecordsetOpen(GDb, wRs, wstr)
            Do Until wRs.eof
                wリース = MBA010_リースデータセット(wRs)
         
                '** リーステーブル セット **
                If wリース.手入力区分 <> P8.FCDbl(XMXA020_区分("登録方法", "標準登録")) Then                  ' 07/02/12 V180
                    Call MBA010_科目集計セット_リース手入力(p更新モード, wリース, p売上計画) ' 07/02/12 V180
                Else                                            ' 07/02/12 V180
                    Call MBA010_リーステーブル作成(wリース)

                    '** 科目テーブル セット **
                    Call MBA010_科目集計セット_リース(p更新モード, wリース, p売上計画)
                End If                                          ' 07/02/12 V180
                
                wRs.MoveNext
            Loop
        wRs.Close
        Set wRs = Nothing
        
        
    End If
    
    
    
    
    
    ' -----------------------------------------
    '     科目テーブル 縦計算
    ' -----------------------------------------
    Call MBA010_テーブル縦計算(p更新モード, p売上計画, w基本事業計画, w対象年次, pボトムアップ) '06/09/21 V170
    
 '   Call MAA500_科目デバッグ
    
    ' -----------------------------------------
    '     科目集計テーブル 作成
    ' -----------------------------------------
    If p更新モード = "損益予実対比表予算" Or p更新モード = "経営計画支援表２" Then
        wFILE名 = "DECA020_科目集計2"
        wFILE名2 = "DXAA020_科目テーブルデバック２"
    Else
        wFILE名 = "DECA010_科目集計"
        wFILE名2 = "DXAA020_科目テーブルデバック"
    End If
    
    If p更新モード = "損益予実対比表予算" Or p借入計画番号 = "" Then
        If p更新モード = "損益資金計画推移表" Or p更新モード = "損益計画推移表" _
            Or p更新モード = "損益予実対比表" Or p更新モード = "損益予実対比表予算" _
            Or p更新モード = "経営計画支援表" Or p更新モード = "経営計画支援表２" _
            Or p更新モード = "企業継続分岐点売上表" Or p更新モード = "損益分岐点売上表" Then
            
            Call MCA010_科目作成Write(p更新モード, p売上計画.売上計画開始年度, _
                          p推移表開始年度, wFILE名, _
                          p推移表区分)
            
            Call MAA500_科目デバッグ(wFILE名2)   '2004/10/6 V120
            
            Exit Sub
        End If
    End If
    
   
    ' -----------------------------------------
    '     科目テーブル  資金調達計算
    ' -----------------------------------------
    w対象年次 = 0
   
    wRet資金調達 = MBA010_資金調達(p売上計画, w対象年次, w対象最終年次)
    
    
    '***資産売却に伴う処置
    G仮払消費税22F = 1                  ' 07/03/02 V180
     
    ' -----------------------------------------
    ' 出力用借入金テーブル セット ＆ 次の借入計算
    ' -----------------------------------------
    w配列数 = 0
    ReDim w出力用借入金テーブル(w配列数)
    
    While (Not wRet資金調達.終了)
        
        '** 出力用借入金テーブル セット **
        w配列数 = w配列数 + 1
        ReDim Preserve w出力用借入金テーブル(w配列数)
     
        w借入金 = wRet資金調達.借入金データ
        w出力用借入金テーブル(w配列数) = w借入金
         
        '** 出力用借入金を元に再度計算 **
        Call MBD010_借入金テーブル作成(p金融リストラ, w借入金)
        
        'Call MBA010_科目集計セット_借入金(w借入金, p売上計画, p金融リストラ)
        'Call MBA010_テーブル縦計算(p売上計画, w基本事業計画)
        'wRet資金調達 = MBA010_資金調達(p売上計画)
        wRet資金調達.終了 = MBA010_科目集計セット_借入金(p更新モード, w借入金, p売上計画, p金融リストラ)
        If Not wRet資金調達.終了 Then
            Call MBA010_テーブル縦計算(p更新モード, p売上計画, w基本事業計画, w対象年次, pボトムアップ) '06/09/21 V170
            wRet資金調達 = MBA010_資金調達(p売上計画, w対象年次, w対象最終年次)
        End If
    Wend
    
      
    ' -----------------------------------------
    ' 出力用借入金テーブル を データベースに出力
    ' -----------------------------------------
    w自動借入連番 = 0
    wstr1 = ""
    wstr1 = wstr1 + "Select * From DAAA020_コントロール"
    wstr1 = wstr1 + " Where System = 'System'"
    Call AdoRecordsetOpen(GDb, wRs1, wstr1)
        If Not wRs1.eof Then
            If w自動借入連番 > 999999 Then
                w自動借入連番 = 0
            Else
                w自動借入連番 = wRs1("自動借入連番")
            End If
        End If
    wRs1.Close
    Set wRs1 = Nothing
    
    w借入番号 = ""
    wstr = ""
    wstr = wstr + "Select * From DBDA010_借入金"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        For j = 1 To w配列数
            wRs.AddNew
            
            w自動借入連番 = w自動借入連番 + 1
            
            wRs("借入番号") = w出力用借入金テーブル(j).借入番号 & w自動借入連番
            wRs("手入力区分") = w出力用借入金テーブル(j).手入力区分     ' 07/02/12 V180
              
            wRs("借入計画番号") = w出力用借入金テーブル(j).借入計画番号
            wRs("借入内容") = w出力用借入金テーブル(j).借入内容
            wRs("Sm区分") = w出力用借入金テーブル(j).SM区分
            wRs("金融リストラ番号") = w出力用借入金テーブル(j).金融リストラ番号
            'wRs("金融リストラ区分") = w出力用借入金テーブル(J).金融リストラ区分
            wRs("銀行番号") = w出力用借入金テーブル(j).銀行番号
            wRs("支払日") = w出力用借入金テーブル(j).支払日             ' 07/02/12 V180
            wRs("営業日区分") = w出力用借入金テーブル(j).営業日区分     ' 07/02/12 V180
            wRs("利息区分") = w出力用借入金テーブル(j).利息区分         ' 07/02/12 V180
            wRs("利息計算日数区分") = w出力用借入金テーブル(j).利息計算日数区分 ' 07/02/12 V180
            wRs("利息支払方法") = w出力用借入金テーブル(j).利息支払方法 ' 07/02/12 V180
            wRs("利息控除区分") = w出力用借入金テーブル(j).利息控除区分 ' 07/02/12 V180
            wRs("金利計算年間日数") = w出力用借入金テーブル(j).金利計算年間日数 ' 07/02/12 V180
            wRs("金利初回年月") = w出力用借入金テーブル(j).金利初回年月 ' 07/02/12 V180
            wRs("融資金額") = w出力用借入金テーブル(j).融資金額
            wRs("利率") = w出力用借入金テーブル(j).利率
            wRs("保証料率") = w出力用借入金テーブル(j).保証料率
            wRs("保証料分割フラグ") = w出力用借入金テーブル(j).保証料分割フラグ
            wRs("実行日") = w出力用借入金テーブル(j).実行日
            wRs("初回返済年月") = w出力用借入金テーブル(j).初回返済年月
            wRs("初回返済実行日") = w出力用借入金テーブル(j).初回返済実行日
            wRs("最終返済年月") = w出力用借入金テーブル(j).最終返済年月
            wRs("最終返済実行日") = w出力用借入金テーブル(j).最終返済実行日
            wRs("解約年月") = w出力用借入金テーブル(j).解約年月
            wRs("解約実行日") = w出力用借入金テーブル(j).解約実行日
            wRs("解約保証料戻") = w出力用借入金テーブル(j).解約保証料戻
            wRs("金融解約年月") = w出力用借入金テーブル(j).金融解約年月
            wRs("金融解約実行日") = w出力用借入金テーブル(j).金融解約実行日
            wRs("金融解約保証料戻") = w出力用借入金テーブル(j).金融解約保証料戻
            wRs("初回返済額") = w出力用借入金テーブル(j).初回返済額
            wRs("毎月返済額") = w出力用借入金テーブル(j).毎月返済額
            wRs("最終返済額") = w出力用借入金テーブル(j).最終返済額
            wRs("返済単位月数") = w出力用借入金テーブル(j).返済単位月数
            wRs("有担保フラグ") = w出力用借入金テーブル(j).有担保フラグ
            wRs("設備フラグ") = w出力用借入金テーブル(j).設備フラグ
            wRs("自己資金フラグ") = w出力用借入金テーブル(j).自己資金フラグ
            wRs("支払回数") = w出力用借入金テーブル(j).支払回数
            wRs("据置回数") = w出力用借入金テーブル(j).据置回数
            wRs("借入貸付") = w出力用借入金テーブル(j).借入貸付     '06/03/11 V150
            wRs("返済方法") = w出力用借入金テーブル(j).返済方法     '06/04/02 V150
            
            For k = 2 To 15
                wRs("金利変更" + CStr(k) + "回目年月") = w出力用借入金テーブル(j).金利(k).金利変更x回目年月
                wRs("金利" + CStr(k) + "回目") = w出力用借入金テーブル(j).金利(k).金利x回目
            Next
                
            wRs("融資可能枠") = w出力用借入金テーブル(j).融資可能枠
            wRs("融資残高") = w出力用借入金テーブル(j).融資残高
            wRs("借入年度") = w出力用借入金テーブル(j).借入年度
            wRs("取消フラグ") = w出力用借入金テーブル(j).取消フラグ
    
            wRs.Update
        Next
    wRs.Close
    Set wRs = Nothing
    
    wstr1 = ""
    wstr1 = wstr1 + "Select * From DAAA020_コントロール"
    wstr1 = wstr1 + " Where System = 'System'"
    Call AdoRecordsetOpen(GDb, wRs1, wstr1)
        If Not wRs1.eof Then
            wRs1("自動借入連番") = w自動借入連番
            wRs1.Update
        End If
    wRs1.Close
    Set wRs1 = Nothing
    
    ' -----------------------------------------
    '     科目集計テーブル 作成 借入計画番号<>""
    ' -----------------------------------------
   
    ' -----------------------------------------
    '     科目集計テーブル 作成
    ' -----------------------------------------
    If p更新モード = "損益予実対比表予算" Or p更新モード = "経営計画支援表２" Then
        wFILE名 = "DECA020_科目集計2"
        wFILE名2 = "DXAA020_科目テーブルデバック２"
    Else
        wFILE名 = "DECA010_科目集計"
        wFILE名2 = "DXAA020_科目テーブルデバック"
    End If
    
    If G会議 = "本部予算" _
    And (wFILE名 = "DECA010_科目集計" Or wFILE名2 = "DXAA020_科目テーブルデバック") Then
    
    Else
    
        If p更新モード = "損益資金計画推移表" Or p更新モード = "損益計画推移表" _
            Or p更新モード = "損益予実対比表" Or p更新モード = "損益予実対比表予算" _
            Or p更新モード = "経営計画支援表" Or p更新モード = "経営計画支援表２" _
            Or p更新モード = "企業継続分岐点売上表" Or p更新モード = "損益分岐点売上表" Then
            
            Call MCA010_科目作成Write(p更新モード, p売上計画.売上計画開始年度, _
                              p推移表開始年度, _
                              wFILE名, _
                              p推移表区分)
                              
            Call MAA500_科目デバッグ(wFILE名2)   '2004/10/6 V120
            
        End If
        
    End If
    
    
    'Call MAA500_科目デバッグ   '2004/10/5
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD020_借入ファイル更新_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD020_借入ファイル更新() でエラー" + vbCrLf + vbCrLf + _
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
' MBD020_売上詳細卸業取消
'------------------------------------------------
Public Sub MBD020_売上詳細卸業取消()
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    Dim w整備売上計画番号(1500) As String
    Dim w整備cnt As Integer
    
    Dim w分類売上計画番号(1500) As String
    Dim w分類(1500) As Integer
    Dim w分類cnt As Integer
    
    Dim w取消フラグ2 As Integer
    
    Dim wcnt As Integer
    Dim w売上計画番号 As String
'
    On Error GoTo MBD020_売上詳細卸業取消_ERR
'
    Erase w整備売上計画番号()
    Erase w分類売上計画番号()
    Erase w分類()
'
    '****
    '  　売上詳細卸業　取消フラグセット
    '****
    w整備cnt = 1
    wstr = ""
    wstr = wstr + "SELECT * FROM DBAA060_売上詳細卸業"
    wstr = wstr + " Where 売上計画取消フラグ = 1"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
        
            w整備売上計画番号(w整備cnt) = wRs("売上計画番号")
            w整備cnt = w整備cnt + 1
                                  
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
    
    w整備売上計画番号(w整備cnt) = "??????????"
    
    
    '**売上詳細卸業マスタ売上計画取消フラグ設定
    wstr = ""
    wstr = wstr + "SELECT * FROM DBAA060_売上詳細卸業"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
        
            w売上計画番号 = wRs("売上計画番号")
            For wcnt = 1 To w整備cnt
                
                If w整備売上計画番号(wcnt) = Left$(w売上計画番号, 4) _
                    Or w整備売上計画番号(wcnt) = wRs("売上計画番号") Then
                    wRs("売上計画取消フラグ") = 1
                    Exit For
                End If
                
            Next
            
            wRs.Update
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
    
    
    '**取消分類売上計画番号セット**
    w分類cnt = 1
    wstr = ""
    wstr = wstr + "SELECT * FROM DBAA060_売上詳細卸業"
    wstr = wstr + " Where 分類取消フラグ = 1"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
        
            w分類売上計画番号(w分類cnt) = wRs("売上計画番号")
            w分類(w分類cnt) = wRs("分類")
            w分類cnt = w分類cnt + 1
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
    
    w分類売上計画番号(w分類cnt) = "??????????"
    
    '**取消フラグ　セット
    wstr = ""
    wstr = wstr + "SELECT * FROM DBAA060_売上詳細卸業B"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            w売上計画番号 = wRs("売上計画番号")
            w取消フラグ2 = 0
            For wcnt = 1 To w分類cnt
                If w分類売上計画番号(wcnt) = w売上計画番号 And _
                    w分類(wcnt) = wRs("分類") Then
                    w取消フラグ2 = 1
                    Exit For
                End If
            Next
            
            For wcnt = 1 To w整備cnt
                If w整備売上計画番号(wcnt) = Left$(w売上計画番号, 4) _
                    Or w整備売上計画番号(wcnt) = wRs("売上計画番号") Then
                    w取消フラグ2 = 1
                    Exit For
                End If
            Next
            
            wRs("取消フラグ2") = w取消フラグ2
            wRs.Update
        
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    Erase w整備売上計画番号()
    Erase w分類売上計画番号()
    Erase w分類()
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD020_売上詳細卸業取消_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD020_売上詳細卸業取消() でエラー" + vbCrLf + vbCrLf + _
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
' MBD020_売上詳細製造業取消
'------------------------------------------------
Public Sub MBD020_売上詳細製造業取消()
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
    Dim w整備売上計画番号(1500) As String
    Dim w整備cnt As Integer
    
    Dim w分類売上計画番号(1500) As String
    Dim w分類(1500) As Integer
    Dim w分類cnt As Integer
    
    Dim w取消フラグ2 As Integer
    
    Dim wcnt As Integer
    Dim w売上計画番号 As String
'
    On Error GoTo MBD020_売上詳細製造業取消_ERR
'
    Erase w整備売上計画番号()
    Erase w分類売上計画番号()
    Erase w分類()
'
    '****
    '  　売上詳細卸業　取消フラグセット
    '****
    w整備cnt = 1
    wstr = ""
    wstr = wstr + "SELECT * FROM DBAA060_売上詳細製造業"
    wstr = wstr + " Where 売上計画取消フラグ = 1"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
        
            w整備売上計画番号(w整備cnt) = wRs("売上計画番号")
            w整備cnt = w整備cnt + 1
                                  
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
    
    w整備売上計画番号(w整備cnt) = "??????????"
    
    
    '**売上詳細卸業マスタ売上計画取消フラグ設定
    wstr = ""
    wstr = wstr + "SELECT * FROM DBAA060_売上詳細製造業"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
        
            w売上計画番号 = wRs("売上計画番号")
            For wcnt = 1 To w整備cnt
                
                If w整備売上計画番号(wcnt) = Left$(w売上計画番号, 4) _
                    Or w整備売上計画番号(wcnt) = wRs("売上計画番号") Then
                    wRs("売上計画取消フラグ") = 1
                    Exit For
                End If
                
            Next
            
            wRs.Update
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
    
    
    '**取消分類売上計画番号セット**
    w分類cnt = 1
    wstr = ""
    wstr = wstr + "SELECT * FROM DBAA060_売上詳細製造業"
    wstr = wstr + " Where 分類取消フラグ = 1"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
        
            w分類売上計画番号(w分類cnt) = wRs("売上計画番号")
            w分類(w分類cnt) = wRs("分類")
            w分類cnt = w分類cnt + 1
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
    
    w分類売上計画番号(w分類cnt) = "??????????"
    
    '**取消フラグ　セット
    wstr = ""
    wstr = wstr + "SELECT * FROM DBAA060_売上詳細製造業B"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            w売上計画番号 = wRs("売上計画番号")
            w取消フラグ2 = 0
            For wcnt = 1 To w分類cnt
                If w分類売上計画番号(wcnt) = w売上計画番号 And _
                    w分類(wcnt) = wRs("分類") Then
                    w取消フラグ2 = 1
                    Exit For
                End If
            Next
            
            For wcnt = 1 To w整備cnt
                If w整備売上計画番号(wcnt) = Left$(w売上計画番号, 4) _
                    Or w整備売上計画番号(wcnt) = wRs("売上計画番号") Then
                    w取消フラグ2 = 1
                    Exit For
                End If
            Next
            
            wRs("取消フラグ2") = w取消フラグ2
            wRs.Update
        
            wRs.MoveNext
        Loop
    wRs.Close
    Set wRs = Nothing
'
    Erase w整備売上計画番号()
    Erase w分類売上計画番号()
    Erase w分類()
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD020_売上詳細製造業取消_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD020_売上詳細製造業取消() でエラー" + vbCrLf + vbCrLf + _
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
' MBD020_借入金ワークテーブル作成
'------------------------------------------------
Public Sub MBD020_借入金ワークテーブル作成(pTbl As String)
'
    On Error GoTo MBD020_借入金ワークテーブル作成_ERR
'
    Dim wRs As ADODB.Recordset, wRs2 As ADODB.Recordset
    Dim wstr As String, wstr2 As String
    
    Dim j As Integer
    Dim wsTbl As String, wsTbl2 As String

'    Dim wdSta As Date, wdEnd As Date
'
'    wdSta = C年月日.年度開始年月日(GRpt.テキスト_01, "平成")
'
'    Select Case GRpt.推移
'        Case "月次"     '1年間
'            wdEnd = DateAdd("yyyy", 1, wdSta)
'        Case "四半期"   '3年間
'            wdEnd = DateAdd("yyyy", 3, wdSta)
'        Case "半期"     '5年間
'            wdEnd = DateAdd("yyyy", 5, wdSta)
'        Case "年次"     '10年間
'            wdEnd = DateAdd("yyyy", 10, wdSta)
'    End Select
'
    wstr2 = ""
    wstr2 = wstr2 & "Delete * From DCIA010_借入金ワーク"
    GDb.Execute wstr2
'
    wsTbl = "DBDA010_借入金"
    wsTbl2 = "DBDA010_分岐点借入金"
    If pTbl = "DBDA010_借入金" Then
        wsTbl = "DBDA010_借入金"
        wsTbl2 = "DBDA010_分岐点借入金"
    ElseIf pTbl = "DBDA010_貸付金" Then
        wsTbl = "DBDA010_貸付金"
        wsTbl2 = "DBDA010_分岐点貸付金"
    End If
'
    wstr2 = "Select * From DCIA010_借入金ワーク"
    Call AdoRecordsetOpen(GDb, wRs2, wstr2)
    
        wstr = "Select K.* FROM ((" & wsTbl & " As K"
        wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
        wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分)"
        wstr = wstr & " LEFT JOIN DAAA200_部門マスタ As B"
        wstr = wstr & " ON K.プロジェクト番号 = B.部門番号)"
        wstr = wstr & " LEFT JOIN DAAA040_銀行マスタ As G"
        wstr = wstr & " ON K.銀行番号 = G.銀行番号"
        
        wstr = wstr & " Where K.取消フラグ=0"
        
        '利子補給金
        If GRpt.帳票名 = "利息明細表" _
        Or GRpt.帳票名 = "平均金利平均残高推移表" _
        Or GRpt.帳票名 = "平均金利平均残高表" _
        Or GRpt.帳票名 = "借入金時価評価明細表" _
        Or GRpt.帳票名 = "借入金時価評価一覧表" _
        Or GRpt.帳票名 = "借入金時価評価適用金利一覧" _
        Or GRpt.帳票名 = "借入金時価評価一覧表_前期末" _
        Or GRpt.帳票名 = "借入金時価評価一覧表_増減" _
        Or GRpt.帳票名 = "借入金時価評価適用金利一覧" _
        Or GRpt.帳票名 = "借入金時価評価適用金利一覧_前期末" _
        Or GRpt.帳票名 = "仕訳表 -月次処理-" _
        Or GRpt.帳票名 = "仕訳表 -決算処理-" _
        Or GRpt.帳票名 = "金融機関別残高表" _
        Or GRpt.帳票名 = "年度別比較表" _
        Or GRpt.帳票名 = "1年以内返済長期借入金集計表" _
        Or GRpt.帳票名 = "銀行別利息表" _
        Or GRpt.帳票名 = "支払利息推移表" _
        Then
            wstr = wstr & " And S.利子補給金フラグ=0"
        End If
        
        If GRpt.C_種別 <> "" Then
            wstr = wstr & " And S.借入金種別名='" & GRpt.C_種別 & "'"
        End If
        If GRpt.C_部門 <> "" Then
            wstr = wstr & " And B.部門名='" & GRpt.C_部門 & "'"
        End If
        If GRpt.C_金融 <> "" Then
            wstr = wstr & " And G.金融機関名='" & GRpt.C_金融 & "'"
        End If
        If GRpt.C_銀行 <> "" Then
            wstr = wstr & " And G.銀行名='" & GRpt.C_銀行 & "'"
        End If
        
        If GRpt.借入 = "" And GRpt.金融 = "" Then
            wstr = wstr & " And (借入計画番号 is null or 借入計画番号='')"
            wstr = wstr & " And sm区分=0 "
        End If

        Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            wRs2.AddNew
                
                wRs2("借入番号") = wRs("借入番号")
                wRs2("プロジェクト番号") = wRs("プロジェクト番号")
                wRs2("保証会社区分") = wRs("保証会社区分")
                wRs2("融資区分") = wRs("融資区分")
                wRs2("手入力区分") = wRs("手入力区分")
                wRs2("日割計算区分") = wRs("日割計算区分")
                wRs2("借入内容") = wRs("借入内容")
                wRs2("借入計画番号") = wRs("借入計画番号")
                wRs2("sm区分") = wRs("sm区分")
                wRs2("金融リストラ番号") = wRs("金融リストラ番号")
                wRs2("銀行番号") = wRs("銀行番号")
                wRs2("支払日") = wRs("支払日")
                wRs2("営業日区分") = wRs("営業日区分")
                wRs2("利息区分") = wRs("利息区分")
                wRs2("利息計算日数区分") = wRs("利息計算日数区分")
                wRs2("利息支払方法") = wRs("利息支払方法")
                wRs2("利息控除区分") = wRs("利息控除区分")
                wRs2("金利計算年間日数") = wRs("金利計算年間日数")
                wRs2("融資金額") = wRs("融資金額")
                wRs2("利率") = wRs("利率")
                wRs2("保証料率") = wRs("保証料率")
                wRs2("保証料分割フラグ") = wRs("保証料分割フラグ")
                wRs2("実行日") = wRs("実行日")
                wRs2("初回返済年月") = wRs("初回返済年月")
                wRs2("初回返済実行日") = wRs("初回返済実行日")
                wRs2("金利初回年月") = wRs("金利初回年月")
                wRs2("最終返済年月") = wRs("最終返済年月")
                wRs2("最終返済実行日") = wRs("最終返済実行日")
                wRs2("解約年月") = wRs("解約年月")
                wRs2("解約実行日") = wRs("解約実行日")
                wRs2("解約保証料戻") = wRs("解約保証料戻")
                wRs2("金融解約年月") = wRs("金融解約年月")
                wRs2("金融解約実行日") = wRs("金融解約実行日")
                wRs2("金融解約保証料戻") = wRs("金融解約保証料戻")
                wRs2("返済方法") = wRs("返済方法")
                wRs2("借入貸付") = wRs("借入貸付")
                wRs2("借入金種別区分") = wRs("借入金種別区分")
                wRs2("初回返済額") = wRs("初回返済額")
                wRs2("毎月返済額") = wRs("毎月返済額")
                wRs2("最終返済額") = wRs("最終返済額")
                wRs2("返済単位月数") = wRs("返済単位月数")
                wRs2("有担保フラグ") = wRs("有担保フラグ")
                wRs2("担保名") = wRs("担保名")
                wRs2("設備フラグ") = wRs("設備フラグ")
                wRs2("資金用途") = wRs("資金用途")
                wRs2("自己資金フラグ") = wRs("自己資金フラグ")
                wRs2("長短区分") = wRs("長短区分")
                wRs2("支払回数") = wRs("支払回数")
                wRs2("据置回数") = wRs("据置回数")
                wRs2("金利種別") = wRs("金利種別")
                wRs2("金利条件") = wRs("金利条件")
                wRs2("基準金利区分") = wRs("基準金利区分")
                wRs2("金利グループ区分") = wRs("金利グループ区分")
                wRs2("融資可能枠") = wRs("融資可能枠")
                wRs2("融資残高") = wRs("融資残高")
                wRs2("借入年度") = wRs("借入年度")
                wRs2("取消フラグ") = wRs("取消フラグ")

                For j = 2 To 100
                    wRs2("金利変更" & CStr(j) & "回目年月") = wRs("金利変更" & CStr(j) & "回目年月")
                    wRs2("金利" & CStr(j) & "回目") = wRs("金利" & CStr(j) & "回目")
                Next j
                
            wRs2.Update
        
            wRs.MoveNext
        Loop
        wRs.Close
        Set wRs = Nothing
'
        wstr = "Select K.* FROM ((" & wsTbl2 & " As K"
        wstr = wstr & " LEFT JOIN DAAA116_借入金種別 As S"
        wstr = wstr & " ON K.借入金種別区分 = S.借入金種別区分)"
        wstr = wstr & " LEFT JOIN DAAA200_部門マスタ As B"
        wstr = wstr & " ON K.プロジェクト番号 = B.部門番号)"
        wstr = wstr & " LEFT JOIN DAAA040_銀行マスタ As G"
        wstr = wstr & " ON K.銀行番号 = G.銀行番号"
        
        wstr = wstr & " Where K.取消フラグ=0"
        If GRpt.C_種別 <> "" Then
            wstr = wstr & " And S.借入金種別名='" & GRpt.C_種別 & "'"
        End If
        If GRpt.C_部門 <> "" Then
            wstr = wstr & " And B.部門名='" & GRpt.C_部門 & "'"
        End If
        If GRpt.C_金融 <> "" Then
            wstr = wstr & " And G.金融機関名='" & GRpt.C_金融 & "'"
        End If
        If GRpt.C_銀行 <> "" Then
            wstr = wstr & " And G.銀行名='" & GRpt.C_銀行 & "'"
        End If
        
        If GRpt.借入 = "" And GRpt.金融 = "" Then
            wstr = wstr & " And (借入計画番号 is null or 借入計画番号='')"
            wstr = wstr & " And sm区分=0 "
        End If

        Call AdoRecordsetOpen(GDb, wRs, wstr)
        Do Until wRs.eof
            wRs2.AddNew
                
                wRs2("借入番号") = wRs("借入番号")
                wRs2("プロジェクト番号") = wRs("プロジェクト番号")
                wRs2("保証会社区分") = wRs("保証会社区分")
                wRs2("融資区分") = wRs("融資区分")
                wRs2("手入力区分") = wRs("手入力区分")
                wRs2("日割計算区分") = wRs("日割計算区分")
                wRs2("借入内容") = wRs("借入内容")
                wRs2("借入計画番号") = wRs("借入計画番号")
                wRs2("sm区分") = wRs("sm区分")
                wRs2("金融リストラ番号") = wRs("金融リストラ番号")
                wRs2("銀行番号") = wRs("銀行番号")
                wRs2("支払日") = wRs("支払日")
                wRs2("営業日区分") = wRs("営業日区分")
                wRs2("利息区分") = wRs("利息区分")
                wRs2("利息計算日数区分") = wRs("利息計算日数区分")
                wRs2("利息支払方法") = wRs("利息支払方法")
                wRs2("利息控除区分") = wRs("利息控除区分")
                wRs2("金利計算年間日数") = wRs("金利計算年間日数")
                wRs2("融資金額") = wRs("融資金額")
                wRs2("利率") = wRs("利率")
                wRs2("保証料率") = wRs("保証料率")
                wRs2("保証料分割フラグ") = wRs("保証料分割フラグ")
                wRs2("実行日") = wRs("実行日")
                wRs2("初回返済年月") = wRs("初回返済年月")
                wRs2("初回返済実行日") = wRs("初回返済実行日")
                wRs2("金利初回年月") = wRs("金利初回年月")
                wRs2("最終返済年月") = wRs("最終返済年月")
                wRs2("最終返済実行日") = wRs("最終返済実行日")
                wRs2("解約年月") = wRs("解約年月")
                wRs2("解約実行日") = wRs("解約実行日")
                wRs2("解約保証料戻") = wRs("解約保証料戻")
                wRs2("金融解約年月") = wRs("金融解約年月")
                wRs2("金融解約実行日") = wRs("金融解約実行日")
                wRs2("金融解約保証料戻") = wRs("金融解約保証料戻")
                wRs2("返済方法") = wRs("返済方法")
                wRs2("借入貸付") = wRs("借入貸付")
                wRs2("借入金種別区分") = wRs("借入金種別区分")
                wRs2("初回返済額") = wRs("初回返済額")
                wRs2("毎月返済額") = wRs("毎月返済額")
                wRs2("最終返済額") = wRs("最終返済額")
                wRs2("返済単位月数") = wRs("返済単位月数")
                wRs2("有担保フラグ") = wRs("有担保フラグ")
                wRs2("担保名") = wRs("担保名")
                wRs2("設備フラグ") = wRs("設備フラグ")
                wRs2("資金用途") = wRs("資金用途")
                wRs2("自己資金フラグ") = wRs("自己資金フラグ")
                wRs2("長短区分") = wRs("長短区分")
                wRs2("支払回数") = wRs("支払回数")
                wRs2("据置回数") = wRs("据置回数")
                wRs2("金利種別") = wRs("金利種別")
                wRs2("金利条件") = wRs("金利条件")
                wRs2("基準金利区分") = wRs("基準金利区分")
                wRs2("金利グループ区分") = wRs("金利グループ区分")
                wRs2("融資可能枠") = wRs("融資可能枠")
                wRs2("融資残高") = wRs("融資残高")
                wRs2("借入年度") = wRs("借入年度")
                wRs2("取消フラグ") = wRs("取消フラグ")

                For j = 2 To 100
                    wRs2("金利変更" & CStr(j) & "回目年月") = wRs("金利変更" & CStr(j) & "回目年月")
                    wRs2("金利" & CStr(j) & "回目") = wRs("金利" & CStr(j) & "回目")
                Next j
            
            wRs2.Update
        
            wRs.MoveNext
        Loop
        wRs.Close
        Set wRs = Nothing
        
    wRs2.Close
    Set wRs2 = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >---------------------------------------------------
MBD020_借入金ワークテーブル作成_ERR:
    pERR_MES = pPROGRAM_ID + "/ MBD020_借入金ワークテーブル作成() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)
    
    End
'
End Sub
