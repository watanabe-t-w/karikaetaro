Attribute VB_Name = "MAA010_基本情報"
Option Explicit
'
Private Const pPROGRAM_ID As String = "MAA010_基本情報"

'------------------------< 基本情報ファイル >-------------------------
Type MAA010_基本情報ファイル
    支店コード As String
    支店名 As String
    企業区分 As String
    
'    消費税率 As Double
'    納税消費税率 As Double
'    法人税率 As Double

    決算月 As Integer
    決算締日 As Integer                 '05/08/09 V129
    
    上期 As Integer
    下期 As Integer
    上期納税月 As Integer
    下期納税月 As Integer
    上期賞与 As Integer
    下期賞与 As Integer

    借入金管理区分 As String            '05/08/22 V129
    決算サイクル As Integer
'    減価償却法 As String
    予算設定区分 As String
    決算書参照区分 As String
    減価償却費計上 As String
    業種区分 As String
    資金調達区分 As String              '05/08/01 V128
    現金取引 As String
'    標準償却年数 As Integer
    消費税納税条件 As Integer
    納税回数 As Integer
        
    協力積立金決算月 As Integer         '06/02/01 V150
    協力積立金決済回数 As Integer       '06/02/01 V150
    
    回収有無 As Integer                 '6/3/1 V150
    支払有無 As Integer                 '6/3/1 V150
    
    売上1構成比 As Double
    売上2構成比 As Double
    売上3構成比 As Double
    
    粗利率 As Double
    粗利率1 As Double
    粗利率2 As Double
    粗利率3 As Double
    
    支払サイト As Integer
    
    売上回収サイト As Double
    
    売上回収1サイト As Double
    売上回収1サイト1 As Integer
    売上回収1サイト2 As Integer
    売上回収1サイト3 As Integer
    売上回収1構成比1 As Double
    売上回収1構成比2 As Double
    売上回収1構成比3 As Double
    
    売上回収2サイト As Double
    売上回収2サイト1 As Integer
    売上回収2サイト2 As Integer
    売上回収2サイト3 As Integer
    売上回収2構成比1 As Double
    売上回収2構成比2 As Double
    売上回収2構成比3 As Double
    
    売上回収3サイト As Double
    売上回収3サイト1 As Integer
    売上回収3サイト2 As Integer
    売上回収3サイト3 As Integer
    売上回収3構成比1 As Double
    売上回収3構成比2 As Double
    売上回収3構成比3 As Double
    
    仕入支払サイト As Double
    
    仕入支払1サイト As Double
    仕入支払1サイト1 As Integer
    仕入支払1サイト2 As Integer
    仕入支払1サイト3 As Integer
    仕入支払1構成比1 As Double
    仕入支払1構成比2 As Double
    仕入支払1構成比3 As Double
    
    仕入支払2サイト As Double
    仕入支払2サイト1 As Integer
    仕入支払2サイト2 As Integer
    仕入支払2サイト3 As Integer
    仕入支払2構成比1 As Double
    仕入支払2構成比2 As Double
    仕入支払2構成比3 As Double
    
    仕入支払3サイト As Double
    仕入支払3サイト1 As Integer
    仕入支払3サイト2 As Integer
    仕入支払3サイト3 As Integer
    仕入支払3構成比1 As Double
    仕入支払3構成比2 As Double
    仕入支払3構成比3 As Double
   
    有担銀行 As String
    有担利率 As Double
    有担保証率 As Double
    有担融資年数 As Integer
    無担銀行 As String
    無担利率 As Double
    無担保証率 As Double
    無担融資年数 As Integer
    余裕資金 As Double
    
    有担限度額 As Double
    無担限度額 As Double
    設備限度額 As Double
    運転限度額 As Double
    
    受注1対象月数 As Integer            '05/08/09 V129
    受注2対象月数 As Integer            '05/08/09 V129
    受注3対象月数 As Integer            '05/08/09 V129
    発注1対象月数 As Integer            '05/08/09 V129
    発注2対象月数 As Integer            '05/08/09 V129
    発注3対象月数 As Integer            '05/08/09 V129
    
    売上1不課税構成比 As Double
    売上1課税構成比 As Double
    売上1非課税構成比 As Double
    売上2不課税構成比 As Double
    売上2課税構成比 As Double
    売上2非課税構成比 As Double
    売上3不課税構成比 As Double
    売上3課税構成比 As Double
    売上3非課税構成比 As Double
    仕入1不課税構成比 As Double
    仕入1課税構成比 As Double
    仕入1非課税構成比 As Double
    仕入2不課税構成比 As Double
    仕入2課税構成比 As Double
    仕入2非課税構成比 As Double
    仕入3不課税構成比 As Double
    仕入3課税構成比 As Double
    仕入3非課税構成比 As Double
    
    固定経費不課税構成比 As Double
    固定経費課税構成比 As Double
    固定経費非課税構成比 As Double
    変動経費1不課税構成比 As Double
    変動経費1課税構成比 As Double
    変動経費1非課税構成比 As Double
    変動経費2不課税構成比 As Double
    変動経費2課税構成比 As Double
    変動経費2非課税構成比 As Double
    変動経費3不課税構成比 As Double
    変動経費3課税構成比 As Double
    変動経費3非課税構成比 As Double
    その他経費1不課税構成比 As Double
    その他経費1課税構成比 As Double
    その他経費1非課税構成比 As Double
    '保険積立不課税構成比 As Double
    '保険積立課税構成比 As Double
    '保険積立非課税構成比 As Double
    保険積立不課税構成比 As Double
    保険積立課税構成比 As Double
    保険積立非課税構成比 As Double
    受取リベート不課税構成比 As Double      ' 06/02/01 V150
    受取リベート課税構成比 As Double
    受取リベート非課税構成比 As Double
    支払リベート不課税構成比 As Double      ' 06/02/01 V150
    支払リベート課税構成比 As Double
    支払リベート非課税構成比 As Double
    営業外収益不課税構成比 As Double
    営業外収益課税構成比 As Double
    営業外収益非課税構成比 As Double
    営業外費用不課税構成比 As Double
    営業外費用課税構成比 As Double
    営業外費用非課税構成比 As Double
        
    ' < 06/02/01 V130
    給与総額1構成比 As Double
    給与総額2構成比 As Double
    給与総額3構成比 As Double
    給与総額1サイト As Integer
    給与総額2サイト As Integer
    給与総額3サイト As Integer
    賞与額1構成比 As Double
    賞与額2構成比 As Double
    賞与額3構成比 As Double
    賞与額1サイト As Integer
    賞与額2サイト As Integer
    賞与額3サイト As Integer
    固定経費1構成比 As Double
    固定経費2構成比 As Double
    固定経費3構成比 As Double
    固定経費1サイト As Integer
    固定経費2サイト As Integer
    固定経費3サイト As Integer
    変動経費1の1構成比 As Double
    変動経費1の2構成比 As Double
    変動経費1の3構成比 As Double
    変動経費1の1サイト As Integer
    変動経費1の2サイト As Integer
    変動経費1の3サイト As Integer
    変動経費2の1構成比 As Double
    変動経費2の2構成比 As Double
    変動経費2の3構成比 As Double
    変動経費2の1サイト As Integer
    変動経費2の2サイト As Integer
    変動経費2の3サイト As Integer
    変動経費3の1構成比 As Double
    変動経費3の2構成比 As Double
    変動経費3の3構成比 As Double
    変動経費3の1サイト As Integer
    変動経費3の2サイト As Integer
    変動経費3の3サイト As Integer
    その他経費1の1構成比 As Double
    その他経費1の2構成比 As Double
    その他経費1の3構成比 As Double
    その他経費1の1サイト As Integer
    その他経費1の2サイト As Integer
    その他経費1の3サイト As Integer
    定期積金1構成比 As Double
    定期積金2構成比 As Double
    定期積金3構成比 As Double
    定期積金1サイト As Integer
    定期積金2サイト As Integer
    定期積金3サイト As Integer
    協力積立金1構成比 As Double
    協力積立金2構成比 As Double
    協力積立金3構成比 As Double
    協力積立金1サイト As Integer
    協力積立金2サイト As Integer
    協力積立金3サイト As Integer
    保険積立1構成比 As Double
    保険積立2構成比 As Double
    保険積立3構成比 As Double
    保険積立1サイト As Integer
    保険積立2サイト As Integer
    保険積立3サイト As Integer
    受取リベート1構成比 As Double   ' 06/02/01 V150
    受取リベート2構成比 As Double
    受取リベート3構成比 As Double
    受取リベート1サイト As Integer
    受取リベート2サイト As Integer
    受取リベート3サイト As Integer
    支払リベート1構成比 As Double   ' 06/02/01 V150
    支払リベート2構成比 As Double
    支払リベート3構成比 As Double
    支払リベート1サイト As Integer
    支払リベート2サイト As Integer
    支払リベート3サイト As Integer
    営業外収益1構成比 As Double
    営業外収益2構成比 As Double
    営業外収益3構成比 As Double
    営業外収益1サイト As Integer
    営業外収益2サイト As Integer
    営業外収益3サイト As Integer
    営業外費用1構成比 As Double
    営業外費用2構成比 As Double
    営業外費用3構成比 As Double
    営業外費用1サイト As Integer
    営業外費用2サイト As Integer
    営業外費用3サイト As Integer
    ' > 06/02/01 V130
    日付入力区分 As String  '2011.11.24　追加 by m.mino
    CSV日付書式  As String  '2012.10.23　追加 by k.kunita
End Type

'------------------------------------------------
' MAA010_基本情報ファイル_Read
'------------------------------------------------
Public Sub MAA010_基本情報ファイル_Read()
'
    Dim wRs As ADODB.Recordset
    Dim wstr As String
'
    On Error GoTo MAA010_基本情報ファイル_Read_ERR
'
    wstr = ""
    wstr = wstr + "Select * From DAAA010_基本情報"
    wstr = wstr + " Where System = 'System'"
    Call AdoRecordsetOpen(GDb, wRs, wstr)
        If Not wRs.EOF Then
            G基本情報.決算月 = P8.FCDbl(wRs("決算月"))
            G基本情報.決算締日 = P8.FCDbl(wRs("決算締日"))                 '05/08/17 V129
            G基本情報.上期 = P8.FCDbl(wRs("上期"))
            G基本情報.下期 = P8.FCDbl(wRs("下期"))
            G基本情報.上期納税月 = P8.FCDbl(wRs("上期納税月"))
            G基本情報.下期納税月 = P8.FCDbl(wRs("下期納税月"))
            G基本情報.上期賞与 = P8.FCDbl(wRs("上期賞与"))
            G基本情報.下期賞与 = P8.FCDbl(wRs("下期賞与"))
            
            G基本情報.支店コード = P8.FCStr(wRs("支店コード"))
            G基本情報.支店名 = P8.FCStr(wRs("支店名"))
            G基本情報.企業区分 = P8.FCStr(wRs("企業区分"))
            
            G基本情報.借入金管理区分 = P8.FCDbl(wRs("借入金管理区分"))     '05/08/22 V129
            G基本情報.決算サイクル = P8.FCDbl(wRs("決算サイクル"))
'            G基本情報.減価償却法 = P8.FCDbl(wRs("減価償却法"))
            G基本情報.消費税納税条件 = P8.FCDbl(wRs("消費税納税条件"))
            G基本情報.納税回数 = P8.FCDbl(wRs("納税回数"))
            G基本情報.予算設定区分 = P8.FCDbl(wRs("予算設定区分"))
            G基本情報.決算書参照区分 = P8.FCDbl(wRs("決算書参照区分"))
            G基本情報.減価償却費計上 = P8.FCDbl(wRs("減価償却費計上"))
            G基本情報.業種区分 = P8.FCDbl(wRs("業種区分"))
            G基本情報.資金調達区分 = P8.FCDbl(wRs("資金調達区分"))          '05/08/01 V128
            G基本情報.現金取引 = P8.FCDbl(wRs("現金取引"))
            G基本情報.回収有無 = P8.FCDbl(wRs("回収有無"))
            G基本情報.支払有無 = P8.FCDbl(wRs("支払有無"))
'            G基本情報.消費税率 = P8.FCDbl(wRs("消費税率"))
'            G基本情報.納税消費税率 = P8.FCDbl(wRs("納税消費税率"))
'            G基本情報.法人税率 = P8.FCDbl(wRs("法人税率"))
'            G基本情報.標準償却年数 = P8.FCDbl(wRs("標準償却年数"))

            G基本情報.協力積立金決算月 = P8.FCDbl(wRs("協力積立金決算月"))      '06/02/01 V150
            G基本情報.協力積立金決済回数 = P8.FCDbl(wRs("協力積立金決済回数"))  '06/02/01 V150
            
            G基本情報.売上1構成比 = P8.FCDbl(wRs("売上1構成比"))
            G基本情報.売上2構成比 = P8.FCDbl(wRs("売上2構成比"))
            G基本情報.売上3構成比 = P8.FCDbl(wRs("売上3構成比"))
            
            G基本情報.粗利率 = P8.FCDbl(wRs("粗利率"))
            G基本情報.粗利率1 = P8.FCDbl(wRs("粗利率1"))
            G基本情報.粗利率2 = P8.FCDbl(wRs("粗利率2"))
            G基本情報.粗利率3 = P8.FCDbl(wRs("粗利率3"))
            
            G基本情報.支払サイト = P8.FCDbl(wRs("支払サイト"))
            
            G基本情報.売上回収サイト = P8.FCDbl(wRs("売上回収サイト"))
            G基本情報.売上回収1サイト = P8.FCDbl(wRs("売上回収1サイト"))
            G基本情報.売上回収1サイト1 = P8.FCDbl(wRs("売上回収1サイト1"))
            G基本情報.売上回収1サイト2 = P8.FCDbl(wRs("売上回収1サイト2"))
            G基本情報.売上回収1サイト3 = P8.FCDbl(wRs("売上回収1サイト3"))
            G基本情報.売上回収1構成比1 = P8.FCDbl(wRs("売上回収1構成比1"))
            G基本情報.売上回収1構成比2 = P8.FCDbl(wRs("売上回収1構成比2"))
            G基本情報.売上回収1構成比3 = P8.FCDbl(wRs("売上回収1構成比3"))
            G基本情報.売上回収2サイト = P8.FCDbl(wRs("売上回収2サイト"))
            G基本情報.売上回収2サイト1 = P8.FCDbl(wRs("売上回収2サイト1"))
            G基本情報.売上回収2サイト2 = P8.FCDbl(wRs("売上回収2サイト2"))
            G基本情報.売上回収2サイト3 = P8.FCDbl(wRs("売上回収2サイト3"))
            G基本情報.売上回収2構成比1 = P8.FCDbl(wRs("売上回収2構成比1"))
            G基本情報.売上回収2構成比2 = P8.FCDbl(wRs("売上回収2構成比2"))
            G基本情報.売上回収2構成比3 = P8.FCDbl(wRs("売上回収2構成比3"))
            G基本情報.売上回収3サイト = P8.FCDbl(wRs("売上回収3サイト"))
            G基本情報.売上回収3サイト1 = P8.FCDbl(wRs("売上回収3サイト1"))
            G基本情報.売上回収3サイト2 = P8.FCDbl(wRs("売上回収3サイト2"))
            G基本情報.売上回収3サイト3 = P8.FCDbl(wRs("売上回収3サイト3"))
            G基本情報.売上回収3構成比1 = P8.FCDbl(wRs("売上回収3構成比1"))
            G基本情報.売上回収3構成比2 = P8.FCDbl(wRs("売上回収3構成比2"))
            G基本情報.売上回収3構成比3 = P8.FCDbl(wRs("売上回収3構成比3"))
            
            G基本情報.仕入支払サイト = P8.FCDbl(wRs("仕入支払サイト"))     '05/04/07 V127
            
            G基本情報.仕入支払1サイト = P8.FCDbl(wRs("仕入支払1サイト"))   '05/04/07 V127
            G基本情報.仕入支払1サイト1 = P8.FCDbl(wRs("仕入支払1サイト1")) '05/04/07 V127
            G基本情報.仕入支払1サイト2 = P8.FCDbl(wRs("仕入支払1サイト2")) '05/04/07 V127
            G基本情報.仕入支払1サイト3 = P8.FCDbl(wRs("仕入支払1サイト3")) '05/04/07 V127
            G基本情報.仕入支払1構成比1 = P8.FCDbl(wRs("仕入支払1構成比1")) '05/04/07 V127
            G基本情報.仕入支払1構成比2 = P8.FCDbl(wRs("仕入支払1構成比2")) '05/04/07 V127
            G基本情報.仕入支払1構成比3 = P8.FCDbl(wRs("仕入支払1構成比3")) '05/04/07 V127
            G基本情報.仕入支払2サイト = P8.FCDbl(wRs("仕入支払2サイト"))   '05/04/07 V127
            G基本情報.仕入支払2サイト1 = P8.FCDbl(wRs("仕入支払2サイト1")) '05/04/07 V127
            G基本情報.仕入支払2サイト2 = P8.FCDbl(wRs("仕入支払2サイト2")) '05/04/07 V127
            G基本情報.仕入支払2サイト3 = P8.FCDbl(wRs("仕入支払2サイト3")) '05/04/07 V127
            G基本情報.仕入支払2構成比1 = P8.FCDbl(wRs("仕入支払2構成比1")) '05/04/07 V127
            G基本情報.仕入支払2構成比2 = P8.FCDbl(wRs("仕入支払2構成比2")) '05/04/07 V127
            G基本情報.仕入支払2構成比3 = P8.FCDbl(wRs("仕入支払2構成比3")) '05/04/07 V127
            G基本情報.仕入支払3サイト = P8.FCDbl(wRs("仕入支払3サイト"))   '05/04/07 V127
            G基本情報.仕入支払3サイト1 = P8.FCDbl(wRs("仕入支払3サイト1")) '05/04/07 V127
            G基本情報.仕入支払3サイト2 = P8.FCDbl(wRs("仕入支払3サイト2")) '05/04/07 V127
            G基本情報.仕入支払3サイト3 = P8.FCDbl(wRs("仕入支払3サイト3")) '05/04/07 V127
            G基本情報.仕入支払3構成比1 = P8.FCDbl(wRs("仕入支払3構成比1")) '05/04/07 V127
            G基本情報.仕入支払3構成比2 = P8.FCDbl(wRs("仕入支払3構成比2")) '05/04/07 V127
            G基本情報.仕入支払3構成比3 = P8.FCDbl(wRs("仕入支払3構成比3")) '05/04/07 V127

            G基本情報.有担銀行 = P8.FCStr(wRs("有担銀行"))
            G基本情報.有担利率 = P8.FCDbl(wRs("有担利率"))
            G基本情報.有担保証率 = P8.FCDbl(wRs("有担保証率"))
            G基本情報.有担融資年数 = P8.FCDbl(wRs("有担融資年数"))
            G基本情報.無担銀行 = P8.FCStr(wRs("無担銀行"))
            G基本情報.無担利率 = P8.FCDbl(wRs("無担利率"))
            G基本情報.無担保証率 = P8.FCDbl(wRs("無担保証率"))
            G基本情報.無担融資年数 = P8.FCDbl(wRs("無担融資年数"))
            G基本情報.余裕資金 = P8.FCDbl(wRs("余裕資金"))
            
            G基本情報.有担限度額 = P8.FCDbl(wRs("有担限度額"))
            G基本情報.無担限度額 = P8.FCDbl(wRs("無担限度額"))
            G基本情報.設備限度額 = P8.FCDbl(wRs("設備限度額"))
            G基本情報.運転限度額 = P8.FCDbl(wRs("運転限度額"))
            
            G基本情報.受注1対象月数 = P8.FCDbl(wRs("受注１対象月数"))       '05/09/24 V129
            G基本情報.受注2対象月数 = P8.FCDbl(wRs("受注２対象月数"))       '05/09/24 V129
            G基本情報.受注3対象月数 = P8.FCDbl(wRs("受注３対象月数"))       '05/09/24 V129
            G基本情報.発注1対象月数 = P8.FCDbl(wRs("発注１対象月数"))       '05/09/24 V129
            G基本情報.発注2対象月数 = P8.FCDbl(wRs("発注２対象月数"))       '05/09/24 V129
            G基本情報.発注3対象月数 = P8.FCDbl(wRs("発注３対象月数"))       '05/09/24 V129
            
            G基本情報.売上1不課税構成比 = P8.FCDbl(wRs("売上1不課税構成比"))
            G基本情報.売上1課税構成比 = P8.FCDbl(wRs("売上1課税構成比"))
            G基本情報.売上1非課税構成比 = P8.FCDbl(wRs("売上1非課税構成比"))
            G基本情報.売上2不課税構成比 = P8.FCDbl(wRs("売上2不課税構成比"))
            G基本情報.売上2課税構成比 = P8.FCDbl(wRs("売上2課税構成比"))
            G基本情報.売上2非課税構成比 = P8.FCDbl(wRs("売上2非課税構成比"))
            G基本情報.売上3不課税構成比 = P8.FCDbl(wRs("売上3不課税構成比"))
            G基本情報.売上3課税構成比 = P8.FCDbl(wRs("売上3課税構成比"))
            G基本情報.売上3非課税構成比 = P8.FCDbl(wRs("売上3非課税構成比"))
            G基本情報.仕入1不課税構成比 = P8.FCDbl(wRs("仕入1不課税構成比"))
            G基本情報.仕入1課税構成比 = P8.FCDbl(wRs("仕入1課税構成比"))
            G基本情報.仕入1非課税構成比 = P8.FCDbl(wRs("仕入1非課税構成比"))
            G基本情報.仕入2不課税構成比 = P8.FCDbl(wRs("仕入2不課税構成比"))
            G基本情報.仕入2課税構成比 = P8.FCDbl(wRs("仕入2課税構成比"))
            G基本情報.仕入2非課税構成比 = P8.FCDbl(wRs("仕入2非課税構成比"))
            G基本情報.仕入3不課税構成比 = P8.FCDbl(wRs("仕入3不課税構成比"))
            G基本情報.仕入3課税構成比 = P8.FCDbl(wRs("仕入3課税構成比"))
            G基本情報.仕入3非課税構成比 = P8.FCDbl(wRs("仕入3非課税構成比"))
    
            G基本情報.固定経費不課税構成比 = P8.FCDbl(wRs("固定経費不課税構成比"))
            G基本情報.固定経費課税構成比 = P8.FCDbl(wRs("固定経費課税構成比"))
            G基本情報.固定経費非課税構成比 = P8.FCDbl(wRs("固定経費非課税構成比"))
            G基本情報.変動経費1不課税構成比 = P8.FCDbl(wRs("変動経費1不課税構成比"))
            G基本情報.変動経費1課税構成比 = P8.FCDbl(wRs("変動経費1課税構成比"))
            G基本情報.変動経費1非課税構成比 = P8.FCDbl(wRs("変動経費1非課税構成比"))
            G基本情報.変動経費2不課税構成比 = P8.FCDbl(wRs("変動経費2不課税構成比"))
            G基本情報.変動経費2課税構成比 = P8.FCDbl(wRs("変動経費2課税構成比"))
            G基本情報.変動経費2非課税構成比 = P8.FCDbl(wRs("変動経費2非課税構成比"))
            G基本情報.変動経費3不課税構成比 = P8.FCDbl(wRs("変動経費3不課税構成比"))
            G基本情報.変動経費3課税構成比 = P8.FCDbl(wRs("変動経費3課税構成比"))
            G基本情報.変動経費3非課税構成比 = P8.FCDbl(wRs("変動経費3非課税構成比"))
            G基本情報.その他経費1不課税構成比 = P8.FCDbl(wRs("その他経費1不課税構成比"))
            G基本情報.その他経費1課税構成比 = P8.FCDbl(wRs("その他経費1課税構成比"))
            G基本情報.その他経費1非課税構成比 = P8.FCDbl(wRs("その他経費1非課税構成比"))
            G基本情報.保険積立不課税構成比 = P8.FCDbl(wRs("保険積立不課税構成比"))
            G基本情報.保険積立課税構成比 = P8.FCDbl(wRs("保険積立課税構成比"))
            G基本情報.保険積立非課税構成比 = P8.FCDbl(wRs("保険積立非課税構成比"))
            'G基本情報保険積立不課税構成比 = P8.FCDbl(wRs("保険積立不課税構成比"))
            'G基本情報保険積立課税構成比 = P8.FCDbl(wRs("保険積立課税構成比"))
            'G基本情報保険積立非課税構成比 = P8.FCDbl(wRs("保険積立非課税構成比"))
            G基本情報.受取リベート不課税構成比 = P8.FCDbl(wRs("受取リベート不課税構成比"))  ' 06/02/01 V150
            G基本情報.受取リベート課税構成比 = P8.FCDbl(wRs("受取リベート課税構成比"))
            G基本情報.受取リベート非課税構成比 = P8.FCDbl(wRs("受取リベート非課税構成比"))
            G基本情報.支払リベート不課税構成比 = P8.FCDbl(wRs("支払リベート不課税構成比"))  ' 06/02/01 V150
            G基本情報.支払リベート課税構成比 = P8.FCDbl(wRs("支払リベート課税構成比"))
            G基本情報.支払リベート非課税構成比 = P8.FCDbl(wRs("支払リベート非課税構成比"))
            G基本情報.営業外収益不課税構成比 = P8.FCDbl(wRs("営業外収益不課税構成比"))
            G基本情報.営業外収益課税構成比 = P8.FCDbl(wRs("営業外収益課税構成比"))
            G基本情報.営業外収益非課税構成比 = P8.FCDbl(wRs("営業外収益非課税構成比"))
            G基本情報.営業外費用不課税構成比 = P8.FCDbl(wRs("営業外費用不課税構成比"))
            G基本情報.営業外費用課税構成比 = P8.FCDbl(wRs("営業外費用課税構成比"))
            G基本情報.営業外費用非課税構成比 = P8.FCDbl(wRs("営業外費用非課税構成比"))
        
            ' < 06/02/01 V130
            G基本情報.給与総額1構成比 = P8.FCDbl(wRs("給与総額１構成比"))
            G基本情報.給与総額2構成比 = P8.FCDbl(wRs("給与総額２構成比"))
            G基本情報.給与総額3構成比 = P8.FCDbl(wRs("給与総額３構成比"))
            G基本情報.給与総額1サイト = P8.FCDbl(wRs("給与総額１サイト"))
            G基本情報.給与総額2サイト = P8.FCDbl(wRs("給与総額２サイト"))
            G基本情報.給与総額3サイト = P8.FCDbl(wRs("給与総額３サイト"))
            G基本情報.賞与額1構成比 = P8.FCDbl(wRs("賞与額１構成比"))
            G基本情報.賞与額2構成比 = P8.FCDbl(wRs("賞与額２構成比"))
            G基本情報.賞与額3構成比 = P8.FCDbl(wRs("賞与額３構成比"))
            G基本情報.賞与額1サイト = P8.FCDbl(wRs("賞与額１サイト"))
            G基本情報.賞与額2サイト = P8.FCDbl(wRs("賞与額２サイト"))
            G基本情報.賞与額3サイト = P8.FCDbl(wRs("賞与額３サイト"))
            G基本情報.固定経費1構成比 = P8.FCDbl(wRs("固定経費１構成比"))
            G基本情報.固定経費2構成比 = P8.FCDbl(wRs("固定経費２構成比"))
            G基本情報.固定経費3構成比 = P8.FCDbl(wRs("固定経費３構成比"))
            G基本情報.固定経費1サイト = P8.FCDbl(wRs("固定経費１サイト"))
            G基本情報.固定経費2サイト = P8.FCDbl(wRs("固定経費２サイト"))
            G基本情報.固定経費3サイト = P8.FCDbl(wRs("固定経費３サイト"))
            G基本情報.変動経費1の1構成比 = P8.FCDbl(wRs("変動経費１の１構成比"))
            G基本情報.変動経費1の2構成比 = P8.FCDbl(wRs("変動経費１の２構成比"))
            G基本情報.変動経費1の3構成比 = P8.FCDbl(wRs("変動経費１の３構成比"))
            G基本情報.変動経費1の1サイト = P8.FCDbl(wRs("変動経費１の１サイト"))
            G基本情報.変動経費1の2サイト = P8.FCDbl(wRs("変動経費１の２サイト"))
            G基本情報.変動経費1の3サイト = P8.FCDbl(wRs("変動経費１の３サイト"))
            G基本情報.変動経費2の1構成比 = P8.FCDbl(wRs("変動経費２の１構成比"))
            G基本情報.変動経費2の2構成比 = P8.FCDbl(wRs("変動経費２の２構成比"))
            G基本情報.変動経費2の3構成比 = P8.FCDbl(wRs("変動経費２の３構成比"))
            G基本情報.変動経費2の1サイト = P8.FCDbl(wRs("変動経費２の１サイト"))
            G基本情報.変動経費2の2サイト = P8.FCDbl(wRs("変動経費２の２サイト"))
            G基本情報.変動経費2の3サイト = P8.FCDbl(wRs("変動経費２の３サイト"))
            G基本情報.変動経費3の1構成比 = P8.FCDbl(wRs("変動経費３の１構成比"))
            G基本情報.変動経費3の2構成比 = P8.FCDbl(wRs("変動経費３の２構成比"))
            G基本情報.変動経費3の3構成比 = P8.FCDbl(wRs("変動経費３の３構成比"))
            G基本情報.変動経費3の1サイト = P8.FCDbl(wRs("変動経費３の１サイト"))
            G基本情報.変動経費3の2サイト = P8.FCDbl(wRs("変動経費３の２サイト"))
            G基本情報.変動経費3の3サイト = P8.FCDbl(wRs("変動経費３の３サイト"))
            G基本情報.その他経費1の1構成比 = P8.FCDbl(wRs("その他経費１の１構成比"))
            G基本情報.その他経費1の2構成比 = P8.FCDbl(wRs("その他経費１の２構成比"))
            G基本情報.その他経費1の3構成比 = P8.FCDbl(wRs("その他経費１の３構成比"))
            G基本情報.その他経費1の1サイト = P8.FCDbl(wRs("その他経費１の１サイト"))
            G基本情報.その他経費1の2サイト = P8.FCDbl(wRs("その他経費１の２サイト"))
            G基本情報.その他経費1の3サイト = P8.FCDbl(wRs("その他経費１の３サイト"))
            G基本情報.定期積金1構成比 = P8.FCDbl(wRs("定期積金１構成比"))
            G基本情報.定期積金2構成比 = P8.FCDbl(wRs("定期積金２構成比"))
            G基本情報.定期積金3構成比 = P8.FCDbl(wRs("定期積金３構成比"))
            G基本情報.定期積金1サイト = P8.FCDbl(wRs("定期積金１サイト"))
            G基本情報.定期積金2サイト = P8.FCDbl(wRs("定期積金２サイト"))
            G基本情報.定期積金3サイト = P8.FCDbl(wRs("定期積金３サイト"))
            G基本情報.協力積立金1構成比 = P8.FCDbl(wRs("協力積立金１構成比"))
            G基本情報.協力積立金2構成比 = P8.FCDbl(wRs("協力積立金２構成比"))
            G基本情報.協力積立金3構成比 = P8.FCDbl(wRs("協力積立金３構成比"))
            G基本情報.協力積立金1サイト = P8.FCDbl(wRs("協力積立金１サイト"))
            G基本情報.協力積立金2サイト = P8.FCDbl(wRs("協力積立金２サイト"))
            G基本情報.協力積立金3サイト = P8.FCDbl(wRs("協力積立金３サイト"))
            G基本情報.保険積立1構成比 = P8.FCDbl(wRs("保険積立１構成比"))
            G基本情報.保険積立2構成比 = P8.FCDbl(wRs("保険積立２構成比"))
            G基本情報.保険積立3構成比 = P8.FCDbl(wRs("保険積立３構成比"))
            G基本情報.保険積立1サイト = P8.FCDbl(wRs("保険積立１サイト"))
            G基本情報.保険積立2サイト = P8.FCDbl(wRs("保険積立２サイト"))
            G基本情報.保険積立3サイト = P8.FCDbl(wRs("保険積立３サイト"))
            G基本情報.受取リベート1構成比 = P8.FCDbl(wRs("受取リベート１構成比"))   ' 06/02/01 V150
            G基本情報.受取リベート2構成比 = P8.FCDbl(wRs("受取リベート２構成比"))
            G基本情報.受取リベート3構成比 = P8.FCDbl(wRs("受取リベート３構成比"))
            G基本情報.受取リベート1サイト = P8.FCDbl(wRs("受取リベート１サイト"))
            G基本情報.受取リベート2サイト = P8.FCDbl(wRs("受取リベート２サイト"))
            G基本情報.受取リベート3サイト = P8.FCDbl(wRs("受取リベート３サイト"))
            G基本情報.支払リベート1構成比 = P8.FCDbl(wRs("支払リベート１構成比"))   ' 06/02/01 V150
            G基本情報.支払リベート2構成比 = P8.FCDbl(wRs("支払リベート２構成比"))
            G基本情報.支払リベート3構成比 = P8.FCDbl(wRs("支払リベート３構成比"))
            G基本情報.支払リベート1サイト = P8.FCDbl(wRs("支払リベート１サイト"))
            G基本情報.支払リベート2サイト = P8.FCDbl(wRs("支払リベート２サイト"))
            G基本情報.支払リベート3サイト = P8.FCDbl(wRs("支払リベート３サイト"))
            G基本情報.営業外収益1構成比 = P8.FCDbl(wRs("営業外収益１構成比"))
            G基本情報.営業外収益2構成比 = P8.FCDbl(wRs("営業外収益２構成比"))
            G基本情報.営業外収益3構成比 = P8.FCDbl(wRs("営業外収益３構成比"))
            G基本情報.営業外収益1サイト = P8.FCDbl(wRs("営業外収益１サイト"))
            G基本情報.営業外収益2サイト = P8.FCDbl(wRs("営業外収益２サイト"))
            G基本情報.営業外収益3サイト = P8.FCDbl(wRs("営業外収益３サイト"))
            G基本情報.営業外費用1構成比 = P8.FCDbl(wRs("営業外費用１構成比"))
            G基本情報.営業外費用2構成比 = P8.FCDbl(wRs("営業外費用２構成比"))
            G基本情報.営業外費用3構成比 = P8.FCDbl(wRs("営業外費用３構成比"))
            G基本情報.営業外費用1サイト = P8.FCDbl(wRs("営業外費用１サイト"))
            G基本情報.営業外費用2サイト = P8.FCDbl(wRs("営業外費用２サイト"))
            G基本情報.営業外費用3サイト = P8.FCDbl(wRs("営業外費用３サイト"))
            ' > 06/02/01 V130
            G基本情報.日付入力区分 = P8.FCStr(wRs("日付入力区分"))  '2011.11.24 By m.mino
            G基本情報.CSV日付書式 = P8.FCStr(wRs("CSV日付書式"))  '2012.10.23　追加 by k.kunita
        End If
    wRs.Close
    Set wRs = Nothing
'
    If G基本情報.日付入力区分 = "0" Then
        '和暦入力
        Gfmt年 = "ee年"
        Gfmt年月 = "ee年mm月"
        Gfmt年月日 = "ee年mm月dd日"
    Else
        Gfmt年 = "yyyy"
        Gfmt年月 = "yyyy/mm"
        Gfmt年月日 = "yyyy/mm/dd"
        '西暦入力
    End If
    
    '2012.10.24　追加 by k.kunita
    If G基本情報.CSV日付書式 = "0" Then
        Gfmtcsv年月日 = "yyyymmdd"
    Else
        If G基本情報.CSV日付書式 = "1" Then
            Gfmtcsv年月日 = "yyyy/mm/dd"
        Else
            If G基本情報.CSV日付書式 = "2" Then
                Gfmtcsv年月日 = "yyyy-mm-dd"
            Else
                If G基本情報.CSV日付書式 = "3" Then
                    Gfmtcsv年月日 = "yyyy.mm.dd"
                Else
                    If G基本情報.CSV日付書式 = "4" Then
                        Gfmtcsv年月日 = "yymmdd"
                    Else
                        If G基本情報.CSV日付書式 = "5" Then
                            Gfmtcsv年月日 = "yy/mm/dd"
                        Else
                            If G基本情報.CSV日付書式 = "6" Then
                                Gfmtcsv年月日 = "yy-mm-dd"
                            Else
                                If G基本情報.CSV日付書式 = "7" Then
                                    Gfmtcsv年月日 = "yy.mm.dd"
                                    Else
                                        If G基本情報.CSV日付書式 = "8" Then
                                            Gfmtcsv年月日 = "eemmdd"
                                        Else
                                            If G基本情報.CSV日付書式 = "9" Then
                                                Gfmtcsv年月日 = "ee/mm/dd"
                                            Else
                                                If G基本情報.CSV日付書式 = "10" Then
                                                    Gfmtcsv年月日 = "ee-mm-dd"
                                                Else
                                                    Gfmtcsv年月日 = "ee.mm.dd"
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If G基本情報.CSV日付書式 = "8" Or G基本情報.CSV日付書式 = "9" Or G基本情報.CSV日付書式 = "10" Or G基本情報.CSV日付書式 = "11" Then
        Gfmtcsv年 = "ee年"
        Gfmtcsv年月 = "ee年mm月"
    Else
        Gfmtcsv年 = "yyyy"
        Gfmtcsv年月 = "yyyy/mm"
    End If
'
    Exit Sub
'
'----------< ERROR ROUTINE >--------------------------------------------------------
MAA010_基本情報ファイル_Read_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA010_基本情報ファイル_Read() でエラー" + vbCrLf + vbCrLf + _
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
' MAA010_基本LIST
'------------------------------------------------
Public Sub MAA010_基本LIST()
'
    Dim wDb As New ADODB.Connection
    Dim wstr As String
'
    On Error GoTo MAA010_基本LIST_ERR
'
    '----------< List.mdb Open >----------------------------------------------------
    Call AdoDbOpen("Jet", wDb, GSerDir + "\" + GMain, "", , GPwd)
    
    wstr = "UPDATE DAAA070_企業名マスタ"
    wstr = wstr & " SET"
    wstr = wstr & " 決算月=" & G基本情報.決算月 & ","
    wstr = wstr & " 決算締日=" & G基本情報.決算締日 & ","
    wstr = wstr & " 回収有無=" & G基本情報.回収有無 & ","
    wstr = wstr & " 支払有無=" & G基本情報.支払有無 & ","
    wstr = wstr & " 最終実績年月=#" & Gコントロール.最終実績年月 & "#"
    wstr = wstr + " Where 企業名Key='" + GKeyName + "'"
    
    wDb.Execute wstr

    wDb.Close
    Set wDb = Nothing
'
    Exit Sub
'
'----------< ERROR ROUTINE >--------------------------------------------------------
MAA010_基本LIST_ERR:
    pERR_MES = pPROGRAM_ID + "/ MAA010_基本LIST() でエラー" + vbCrLf + vbCrLf + _
                "エラー番号　　：" + CStr(Err.Number) + vbCrLf + _
                "プロジェクト名：" + Err.Source + vbCrLf + _
                "エラー内容　　：" + Err.Description + vbCrLf + vbCrLf + _
                GProduct + "を終了します"
    pERR_RET = MsgBox(pERR_MES, vbOKOnly + vbCritical, pMSGBOX_TYTLE)
    pERR_RET = PUT_LOG(pERR_MES)

    End
'
End Sub
