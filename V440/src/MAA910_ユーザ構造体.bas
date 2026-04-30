Attribute VB_Name = "MAA910_ユーザ構造体"
Option Explicit

'------------------------< システム >--------------------------------
Type MAA910_システム
    Sys  As String  'システム区分       3:LUFU      8:FULL  1:借入金
    Mem  As String  '単一複数ユーザ区分 1:単一      8:複数
    Sit  As Boolean '支店別管理         0:False     1:True
    Lan  As Boolean 'LAN対応区分        2:False     9:True
    Ker  As Boolean '経理取込有無       0:False     1:True
    Han  As Boolean '販売取込有無       0:False     1:True
End Type
'

'------------------------< 企業情報 >--------------------------------
Type MAA910_企業情報
    企業名Key As String
    支店コード As String
    支店名 As String
    DB名 As String
    親会社名 As String
    企業区分 As String

    決算月 As Integer
    決算締日 As Integer
    回収有無 As Integer
    支払有無 As Integer
    最終実績年月 As Date
End Type
'

'------------------------< 会議計画 >--------------------------------
Type MAA910_会議計画
    連番 As Integer
    親会社名 As String
    作成GR As String
    資金 As String
    連結売上 As String
    内容 As String
        
    売上() As String
    借入() As String
    リス() As String
    設備() As String
    金融() As String
    設備R() As String
End Type
'

'------------------------< 推移表タイトル >--------------------------
Type MAA910_推移表タイトル
    作表区分 As Variant
    X番目年月(12) As String
End Type
'

'------------------------< 借入金 >-------------------------
Type MAA910_金利
    金利変更x回目年月 As Variant
    金利x回目 As Double
End Type

Type MAA910_内入            ' 2008/12/01 V189
    内入x回目年月日 As Variant
    内入金額x回目 As Double
    '毎回支払額x回目 As Double
    '最終支払年月x回目 As Variant
    '最終支払額x回目 As Double
    手数料x回目 As Double
End Type

Type MAA910_借入金内入

    '内入(40) As MAA910_内入             'V189 2008/12/01
    '内入回数
    
    内入区分 As Boolean
    内入(500) As MAA910_内入
    
    取消フラグ As Integer

End Type

Type MAA910_借入金
    借入番号 As String
    
    '借換 V189
    保証会社区分 As String
    融資区分 As String
    
    手入力区分 As Integer               ' 07/01/30 V180
    日割計算区分 As Integer
    
    借入内容 As String
    借入計画番号 As String
    SM区分 As Integer
    金融リストラ番号 As String
    プロジェクト番号 As String
    銀行番号 As String
    支払日 As Integer                   ' 07/01/30 V180
    営業日区分 As Integer               ' 07/01/30 V180
    利息区分 As String                  ' 07/01/30 V180
    利息計算日数区分 As Integer         ' 07/01/30 V180
    利息支払方法 As Integer             ' 07/01/30 V180
    利息控除区分 As Integer             ' 07/01/30 V180
    金利計算年間日数 As Integer         ' 07/01/30 V180
    融資金額 As Double
    利率 As Double
    保証料率 As Double
    保証料分割フラグ As Integer
    実行日 As Variant
    初回返済年月 As Variant
    初回返済実行日 As Variant
    金利初回年月 As Variant             ' 07/01/30 V180
    最終返済年月 As Variant
    最終返済実行日 As Variant
    解約年月 As Variant
    解約実行日 As Variant
    解約保証料戻 As Double
    金融解約年月 As Variant
    金融解約実行日 As Variant
    金融解約保証料戻 As Variant
    返済方法 As Integer                 '06/02/01 V150
    借入貸付 As Integer                 '06/02/01 V150
    借入金種別区分 As String            '10/11/01
    社債フラグ As Integer               '11/06/01
    利子補給金フラグ As Integer         '16/03/29
    
    初回返済額 As Double
    毎月返済額 As Double
    最終返済額 As Double
    返済単位月数 As Integer
    有担保フラグ As Integer
    担保名 As String                    '10/11/01
    金利種別 As Integer
    金利条件 As String
    基準金利区分 As String              '10/11/01
    金利グループ区分 As String          '10/11/01
    長短区分 As Integer                 '10/11/01
    設備フラグ  As Integer
    資金用途 As String                  '10/11/01
    自己資金フラグ As Integer
    支払回数 As Long
    据置回数 As Long
    金利(100) As MAA910_金利
    変動最終利率 As Double              'V182 2008/01/28
    融資可能枠 As Double
    融資残高 As Double
    借入年度 As Integer
    
    取消フラグ As Integer

End Type

Type MAA910_借入金テーブル
    借入番号 As String
    返済回数 As Integer
    据置x回目 As Integer
    返済予定年月 As Variant
    実際年月日 As Variant
    利息計算年月日 As Variant       '10/01/04
    返済金額 As Double
    元金額 As Double
    利息額 As Double
    手数料 As Double                ' 2008/12/01 V189
    保証料 As Double
    金融保証料 As Double
    初期手数料 As Double            ' 2011/06/01
    元金手数料 As Double            ' 2011/06/01
    利息手数料 As Double            ' 2011/06/01
    融資残高 As Double
    日割日数 As Integer
    利息対象期間日数 As Integer     'V182 2008/01/28
    利率 As Double
End Type

Type MAA910_借入金入力
    借入返済年月日 As Date
    利息計算年月日 As Date
    元金 As Double
    利率 As Double
    利息額 As Double
    仮計上利息額 As Double
    返済金額 As Double
    融資残高 As Double
    初期手数料 As Double            ' 2011/06/01
    元金手数料 As Double            ' 2011/06/01
    利息手数料 As Double            ' 2011/06/01
    保証料 As Double
    日割日数 As Integer
    利息対象期間日数 As Integer
End Type

Type MAA910_利息未払前払テーブル
    銀行番号 As String              '11/02/08
    借入番号 As String
    利息区分 As String
    利息計算日数区分 As Integer
    返済年月日 As Date
    締年月 As Date
    月毎NO As Integer
    元金額 As Double
    融資残高 As Double
    利息計算対象額 As Double
    利息額増 As Double
    利息額減 As Double
    利息残高 As Double
    日割日数 As Integer
    利率 As Double
    開始年月日 As Date
    終了年月日 As Date
    据置x回目 As Integer
    利息期間対象日数 As Integer                 '2014/08/26
    利息期間対象額 As Double                    '2014/08/26
    利息調整F As Integer                        '2014/08/26
End Type


'------------------------< 時価評価 >-------------------------
'2015/06/01 V430 時価評価
Type MAA910_時価評価明細
    借入番号 As String
    決算日 As Variant
    適用金利 As Double
    返済年月日 As Variant
    利息計算年月日 As Variant
    元金額 As Double
    利息額 As Double
    返済金額 As Double
    融資残高 As Double
    日割日数 As Integer
    指数 As Double
    分母 As Double
    現価係数 As Double
    現在価値 As Double
End Type

Type MAA910_適用金利
    借入番号 As String
    銀行番号 As String
    基準金利区分 As String
    実行日 As Variant
    融資金額 As Double
    利率 As Double
    最終返済実行日 As Variant
    決算日 As Variant
    借入時長プラ As Double
    借入時PRM As Double
    決算時融資残高 As Double
    決算時長プラ As Double
    決算時適用PRM As Double
    決算時適用金利 As Double
End Type

Type MAA910_時価評価一覧
    借入番号 As String
    銀行番号 As String
    基準金利区分 As String
    実行日 As Variant
    融資金額 As Double
    利率 As Double
    最終返済実行日 As Variant
    決算日 As Variant
    年内元金額 As Double
    年内現在価値 As Double
    年超元金額 As Double
    年超現在価値 As Double
End Type

'------------------------< 借換 >-------------------------
'借換 V189
Type MAA910_借入金借換
    借換計画番号 As String
    年月日 As Variant
    
    借換番号 As String
    借換番号件数 As Integer
    
    w一般融資融資年数 As Integer
    w保証会社融資年数 As Integer
    w融資区分融資年数 As Integer
    wMAX融資区分 As String
    wMAX制度融資年数 As Integer
    w借入返済率 As Double
    
    '現状
    現状件数 As Integer
    借入番号(4) As String
    保証会社区分(4) As String
    融資区分(4) As String
    制度融資区分(4) As Integer
    銀行番号(4) As String
    有担保フラグ(4) As Integer
    残回数(4) As Integer
    残据置(4) As Integer
    利率(4) As Double
    設備フラグ(4)  As Integer
    保証料率(4) As Double
    融資金額(4) As Double
    毎月返済額(4) As Double
    融資残高(4) As Double
    
    借換区分(4) As String
    
    '借換情報
    借換件数 As Integer
    借換借入番号(4) As String
    借換保証会社区分(4) As String
    借換銀行番号(4) As String
    借換有担保フラグ(4) As Integer
    借換融資区分(4) As String
    借換制度融資区分(4) As Integer
    借換利率(4) As Double
    借換保証料率(4) As Double
    借換設備フラグ(4)  As Integer
    借換融資金額(4) As Double
    借換毎月返済額(4) As Double
    借換融資残高(4) As Double
    借換年数(4) As Integer
    借換据置(4) As Integer
    
    代表借入番号(4) As String
    代表借入連番(4) As Integer
End Type

Type MAA910_借入金借換現状
    保証会社区分 As String
    保証会社区分名 As String
    融資区分 As String
    融資区分名 As String
    制度融資区分 As Integer
    銀行番号 As String
    銀行名 As String
    有担保フラグ As Integer
    借入番号 As String
    初回返済年月 As Variant
    最終返済年月 As Variant
    返済単位月数 As Integer
    残回数 As Integer
    残据置 As Integer
    利率 As Double
    設備フラグ  As Integer
    保証料率 As Double
    融資金額 As Double
    毎月返済額 As Double
    融資残高 As Double
End Type
'

'------------------------< リース >-------------------------
Type MAA910_リース
    リース番号 As String
    手入力区分 As Integer
    リース内容 As String
    リース計画番号 As String
    SM区分 As Integer
    リース会社番号 As String
    支払日 As Integer
    営業日区分 As Integer
    リース総額 As Double
    消費税率摘要区分 As Integer
    消費税総額 As Double
    消費税率 As Double
    実行日 As Variant
    初回支払年月 As Variant
    初回支払実行日 As Variant
    最終支払年月 As Variant
    最終支払実行日 As Variant
    解約年月 As Variant
    解約実行日 As Variant
    初回リース料 As Double
    毎月リース料 As Double
    最終リース料 As Double
    据置回数 As Integer
    支払回数 As Long
    取消フラグ As Integer
End Type

Type MAA910_リーステーブル
    リース番号 As String
    支払回数 As Integer
    据置x回目 As Integer
    支払予定年月 As Variant
    実際年月日 As Variant
    支払合計額 As Double
    リース料 As Double
    消費税率 As Double
    消費税額 As Double
    支払残高 As Double
    正味支払残高 As Double
End Type

Type MAA910_リース入力
    リース支払年月日 As Date
    リース料 As Double
    消費税額 As Double
    支払合計額 As Double
    支払残高 As Double
    正味支払残高 As Double
End Type
'

'
'----------< LOGFILE >----------------------------------------------------
Public Type TYPE_SYSLOG
    LOG_DATE As String * 10
    FILLER00 As String * 1
    LOG_TIME As String * 8
    FILLER01 As String * 1
    LOG_MESS As String * 106
    LOG_CR As Byte
    LOG_LF As Byte
End Type
'
Public Type TYPE_SYSTBL
    LOG_DATE As String
    LOG_KUBN As String
    LOG_SHOR As String
    LOG_MYPC As String
    LOG_KKEY As String
    LOG_KNAM As String
    LOG_MESS As String
    LOG_CR As Byte
    LOG_LF As Byte
End Type

'------<帳票構造体>-------------------------------------------2011.11.14 By m.mino

Public Type MAA910_借入明細表照会
    借入番号 As String
    金融リストラ番号 As String
    金融解約日 As String
    金利シミュレーションGP As String
    入力モード As String
End Type
'

'
'------------------------< 科目マスタ >------------------------------
Type MAA910_科目マスタ
    科目番号 As Long
    科目名 As String
    科目残高区分 As Integer
    全ゼロ区分 As Integer
    損益印刷 As Integer
    資金印刷 As Integer
    損益資金印刷 As Integer
    分岐点印刷 As Integer
    
    支店BotUP As Integer     'V150 2006/07/01
    連結BotUP As Integer     'V150 2006/07/01
End Type
'

'------------------------< 基本事業計画 >----------------------------
Type MAA910_基本事業計画
    事業計画開始年度 As Integer
    売上予算X年次(11) As Double                         '06/09/11 V170
    人数X年次(11) As Double                             '06/09/11 V170
    売上指数X月度(12) As Long
    
    売上回収サイト As Double
    
    売上回収1サイト As Double
    売上回収1サイト1 As Double
    売上回収1サイト2 As Double
    売上回収1サイト3 As Double
    売上回収1構成比1 As Double
    売上回収1構成比2 As Double
    売上回収1構成比3 As Double
    
    売上回収2サイト As Double
    売上回収2サイト1 As Double
    売上回収2サイト2 As Double
    売上回収2サイト3 As Double
    売上回収2構成比1 As Double
    売上回収2構成比2 As Double
    売上回収2構成比3 As Double
    
    売上回収3サイト As Double
    売上回収3サイト1 As Double
    売上回収3サイト2 As Double
    売上回収3サイト3 As Double
    売上回収3構成比1 As Double
    売上回収3構成比2 As Double
    売上回収3構成比3 As Double
    
    仕入支払サイト As Double
    
    仕入支払1サイト As Double
    仕入支払1サイト1 As Double
    仕入支払1サイト2 As Double
    仕入支払1サイト3 As Double
    仕入支払1構成比1 As Double
    仕入支払1構成比2 As Double
    仕入支払1構成比3 As Double
    
    仕入支払2サイト As Double
    仕入支払2サイト1 As Double
    仕入支払2サイト2 As Double
    仕入支払2サイト3 As Double
    仕入支払2構成比1 As Double
    仕入支払2構成比2 As Double
    仕入支払2構成比3 As Double
    
    仕入支払3サイト As Double
    仕入支払3サイト1 As Double
    仕入支払3サイト2 As Double
    仕入支払3サイト3 As Double
    仕入支払3構成比1 As Double
    仕入支払3構成比2 As Double
    仕入支払3構成比3 As Double
   
    粗利率 As Double
    粗利率1 As Double
    粗利率2 As Double
    粗利率3 As Double
    
    売上1構成比 As Double
    売上2構成比 As Double
    売上3構成比 As Double
    
    売上達成率 As Double
    売上1達成率 As Double
    売上2達成率 As Double
    売上3達成率 As Double
    
    給与up率 As Double
    賞与up率 As Double
    新人給与月額 As Double
    新人賞与額 As Double
    
    給与総額達成率 As Double
    賞与額達成率 As Double
    固定経費達成率 As Double
    変動経費1達成率 As Double
    変動経費2達成率 As Double
    変動経費3達成率 As Double
    その他経費1達成率 As Double
    定期積金達成率 As Double            '06/02/01 V150
    協力積立金達成率 As Double          '06/02/01 V150
    保険積立達成率 As Double
    受取リベート達成率 As Double        '06/02/01 V150
    支払リベート達成率 As Double        '06/02/01 V150
    営業外収益達成率 As Double
    営業外費用達成率 As Double
    減価償却費達成率 As Double
    支払利息達成率 As Double
    
    給与総額 As Double
    賞与額 As Double
    固定経費 As Double
    変動経費1 As Double
    変動経費2 As Double
    変動経費3 As Double
    その他経費1 As Double
    定期積金 As Double            '06/02/01 V150
    協力積立金 As Double          '06/02/01 V150
    保険積立 As Double
    受取リベート As Double        '06/02/01 V150
    支払リベート As Double        '06/02/01 V150
    営業外収益 As Double
    営業外費用 As Double
    減価償却費 As Double
    支払利息 As Double
    
'    本社非現金費用 As Double      ' V200
'    本社現金費用 As Double        ' V200
    
    手持資金 As Double
    定期積金残 As Double          '06/02/01 V150
    協力積立金残 As Double        '06/02/01 V150
    その他資金1 As Double
    その他資金2 As Double
    その他資金3 As Double
    売掛残高 As Double
    買掛残高 As Double
    投資債権残 As Double          '06/02/01 V150
    投資債務残 As Double          '06/02/01 V150
    未収入金残 As Double          '06/03/11 V150
    未払費用残 As Double          '06/02/01 V150
    期末在庫 As Double            '05/05/05 V127
    前期繰越利益 As Double        '05/08/09 V129
    その他債権残高 As Double      '05/07/23 V128
    その他債権残高2 As Double     '06/02/01 V150
    その他債務残高 As Double      '05/07/23 V128
    その他債務残高2 As Double     '06/02/01 V150
    
    取消フラグ As Integer
End Type
'

'------------------------< 売上計画 >-------------------------
Type MAA910_売上計画
    売上計画番号 As String
    売上計画内容 As String
    設備計画番号 As String
    売上計画開始年度 As Integer
    売上計画開始年月 As Variant
    売上予算前年次 As Double
    人数前年次 As Long
    売上予算X年次(11) As Double                         '06/09/11 V170
    人数X年次(11) As Long                               '06/09/11 V170
    売上指数X月度(12) As Long
    
    売上回収サイト As Double
    
    売上回収1サイト As Double
    売上回収1サイト1 As Double
    売上回収1サイト2 As Double
    売上回収1サイト3 As Double
    売上回収1構成比1 As Double
    売上回収1構成比2 As Double
    売上回収1構成比3 As Double
    
    売上回収2サイト As Double
    売上回収2サイト1 As Double
    売上回収2サイト2 As Double
    売上回収2サイト3 As Double
    売上回収2構成比1 As Double
    売上回収2構成比2 As Double
    売上回収2構成比3 As Double
    
    売上回収3サイト As Double
    売上回収3サイト1 As Double
    売上回収3サイト2 As Double
    売上回収3サイト3 As Double
    売上回収3構成比1 As Double
    売上回収3構成比2 As Double
    売上回収3構成比3 As Double
    
    仕入支払サイト As Double
    
    仕入支払1サイト As Double
    仕入支払1サイト1 As Double
    仕入支払1サイト2 As Double
    仕入支払1サイト3 As Double
    仕入支払1構成比1 As Double
    仕入支払1構成比2 As Double
    仕入支払1構成比3 As Double
    
    仕入支払2サイト As Double
    仕入支払2サイト1 As Double
    仕入支払2サイト2 As Double
    仕入支払2サイト3 As Double
    仕入支払2構成比1 As Double
    仕入支払2構成比2 As Double
    仕入支払2構成比3 As Double
    
    仕入支払3サイト As Double
    仕入支払3サイト1 As Double
    仕入支払3サイト2 As Double
    仕入支払3サイト3 As Double
    仕入支払3構成比1 As Double
    仕入支払3構成比2 As Double
    仕入支払3構成比3 As Double
   
    粗利率 As Double
    粗利率1 As Double
    粗利率2 As Double
    粗利率3 As Double
    
    売上1構成比 As Double
    売上2構成比 As Double
    売上3構成比 As Double
    
    売上達成率 As Double
    売上1達成率 As Double
    売上2達成率 As Double
    売上3達成率 As Double
    
    給与up率 As Double
    賞与up率 As Double
    新人給与月額 As Double
    新人賞与額 As Double
    
    給与総額達成率 As Double
    賞与額達成率 As Double
    固定経費達成率 As Double
    変動経費1達成率 As Double
    変動経費2達成率 As Double
    変動経費3達成率 As Double
    その他経費1達成率 As Double
    定期積金達成率 As Double            '06/02/01 V150
    協力積立金達成率 As Double          '06/02/01 V150
    保険積立達成率 As Double
    受取リベート達成率 As Double        '06/02/01 V150
    支払リベート達成率 As Double        '06/02/01 V150
    営業外収益達成率 As Double
    営業外費用達成率 As Double
    減価償却費達成率 As Double
    支払利息達成率 As Double
    
    給与総額 As Double
    賞与額 As Double
    固定経費 As Double
    変動経費1 As Double
    変動経費2 As Double
    変動経費3 As Double
    その他経費1 As Double
    定期積金 As Double            '06/02/01 V150
    協力積立金 As Double          '06/02/01 V150
    保険積立 As Double
    受取リベート As Double        '06/02/01 V150
    支払リベート As Double        '06/02/01 V150
    営業外収益 As Double
    営業外費用 As Double
    減価償却費 As Double
    支払利息 As Double
    
'    本社非現金費用 As Double      ' V200
'    本社現金費用 As Double        ' V200
    
    手持資金 As Double
    定期積金残 As Double          '06/02/01 V150
    協力積立金残 As Double        '06/02/01 V150
    その他資金1 As Double
    その他資金2 As Double
    その他資金3 As Double
    売掛残高 As Double
    買掛残高 As Double
    投資債権残 As Double          '06/02/01 V150
    投資債務残 As Double          '06/02/01 V150
    未収入金残 As Double          '06/03/11 V150
    未払費用残 As Double          '06/02/01 V150
    期末在庫 As Double            '05/05/05 V127
    前期繰越利益 As Double        '05/08/09 V129
    その他債権残高 As Double      '05/07/23 V128
    その他債権残高2 As Double     '06/02/01 V150
    その他債務残高 As Double      '05/07/23 V128
    その他債務残高2 As Double     '06/02/01 V150
    
    取消フラグ As Integer
End Type
'

'
Type MAA910_売上計画テーブル
    売上計画番号 As String
    売上計画年月 As Variant
    売上計画年度 As Integer
    人数 As Long
    
    粗利率 As Double
    粗利率1 As Double
    粗利率2 As Double
    粗利率3 As Double
    
    売上1構成比 As Double
    売上2構成比 As Double
    売上3構成比 As Double
    
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
    仕入支払1サイト1 As Double
    仕入支払1サイト2 As Double
    仕入支払1サイト3 As Double
    仕入支払1構成比1 As Double
    仕入支払1構成比2 As Double
    仕入支払1構成比3 As Double
    
    仕入支払2サイト As Double
    仕入支払2サイト1 As Double
    仕入支払2サイト2 As Double
    仕入支払2サイト3 As Double
    仕入支払2構成比1 As Double
    仕入支払2構成比2 As Double
    仕入支払2構成比3 As Double
    
    仕入支払3サイト As Double
    仕入支払3サイト1 As Double
    仕入支払3サイト2 As Double
    仕入支払3サイト3 As Double
    仕入支払3構成比1 As Double
    仕入支払3構成比2 As Double
    仕入支払3構成比3 As Double
    
    売上額 As Double
    売上1 As Double
    売上2 As Double
    売上3 As Double
    粗利益 As Double
    粗利1 As Double
    粗利2 As Double
    粗利3 As Double
    
    給与総額 As Double
    賞与額 As Double
    固定経費 As Double
    変動経費1 As Double
    変動経費2 As Double
    変動経費3 As Double
    その他経費1 As Double
    定期積金 As Double                '06/02/01 V150
    協力積立金 As Double              '06/02/01 V150
    保険積立 As Double
    受取リベート As Double                '06/02/01 V150
    支払リベート As Double                '06/02/01 V150
    営業外収益 As Double
    営業外費用 As Double
    減価償却費 As Double
    支払利息 As Double
End Type
'

'------------------------< 年度計画 >------------------------- 06/09/11 V170
Type MAA910_年度計画
    年度(11) As Integer
    
    粗利率(11) As Double
    粗利率1(11) As Double
    粗利率2(11) As Double
    粗利率3(11) As Double
    
    売上1構成比(11) As Double
    売上2構成比(11) As Double
    売上3構成比(11) As Double
    
    換算1構成比(11) As Double
    換算2構成比(11) As Double
    換算3構成比(11) As Double
    
    売上回収サイト(11) As Double
    
    売上回収1サイト(11) As Double
    売上回収1サイト1(11) As Integer
    売上回収1サイト2(11) As Integer
    売上回収1サイト3(11) As Integer
    売上回収1構成比1(11) As Double
    売上回収1構成比2(11) As Double
    売上回収1構成比3(11) As Double
    
    売上回収2サイト(11) As Double
    売上回収2サイト1(11) As Integer
    売上回収2サイト2(11) As Integer
    売上回収2サイト3(11) As Integer
    売上回収2構成比1(11) As Double
    売上回収2構成比2(11) As Double
    売上回収2構成比3(11) As Double
    
    売上回収3サイト(11) As Double
    売上回収3サイト1(11) As Integer
    売上回収3サイト2(11) As Integer
    売上回収3サイト3(11) As Integer
    売上回収3構成比1(11) As Double
    売上回収3構成比2(11) As Double
    売上回収3構成比3(11) As Double
    
    仕入支払サイト(11) As Double
    
    仕入支払1サイト(11) As Double
    仕入支払1サイト1(11) As Integer
    仕入支払1サイト2(11) As Integer
    仕入支払1サイト3(11) As Integer
    仕入支払1構成比1(11) As Double
    仕入支払1構成比2(11) As Double
    仕入支払1構成比3(11) As Double
    
    仕入支払2サイト(11) As Double
    仕入支払2サイト1(11) As Integer
    仕入支払2サイト2(11) As Integer
    仕入支払2サイト3(11) As Integer
    仕入支払2構成比1(11) As Double
    仕入支払2構成比2(11) As Double
    仕入支払2構成比3(11) As Double
    
    仕入支払3サイト(11) As Double
    仕入支払3サイト1(11) As Integer
    仕入支払3サイト2(11) As Integer
    仕入支払3サイト3(11) As Integer
    仕入支払3構成比1(11) As Double
    仕入支払3構成比2(11) As Double
    仕入支払3構成比3(11) As Double
    
    売上予算(11) As Double
    売上予算1(11) As Double
    売上予算2(11) As Double
    売上予算3(11) As Double
    粗利予算(11) As Double
    粗利予算1(11) As Double
    粗利予算2(11) As Double
    粗利予算3(11) As Double
    
    給与総額(11) As Double
    賞与額(11) As Double
    固定経費(11) As Double
    変動経費1(11) As Double
    変動経費2(11) As Double
    変動経費3(11) As Double
    その他経費1(11) As Double
    定期積金(11) As Double            '06/02/01 V150
    協力積立金(11) As Double          '06/02/01 V150
    保険積立(11) As Double
    受取リベート(11) As Double            '06/02/01 V150
    支払リベート(11) As Double            '06/02/01 V150
    営業外収益(11) As Double
    営業外費用(11) As Double
    減価償却費(11) As Double
    支払利息(11) As Double

    給与指数(11) As Double
    賞与指数(11) As Double
    固定経費指数(11) As Double
    変動経費3指数(11) As Double         '06/02/01 V150
    その他経費1指数(11) As Double
    定期積金指数(11) As Double          '06/02/01 V150
    協力積立金指数(11) As Double        '06/02/01 V150
    保険積立指数(11) As Double
    受取リベート指数(11) As Double          '06/02/01 V150
    支払リベート指数(11) As Double          '06/02/01 V150
    営業外収益指数(11) As Double        '06/02/01 V150
    営業外費用指数(11) As Double        '06/02/01 V150
End Type
'
    

'------------------------< 売上実績 >-------------------------
Type MAA910_売上実績
    実績年月 As Date
    売上計画番号 As String
    人数 As Long
    
    売上額 As Double
    売上1 As Double
    売上11 As Double                  '04/08/05  V120
    売上11不課税 As Double            '04/08/05  V120
    売上11課税 As Double              '04/08/05  V120
    売上11非課税 As Double            '04/08/05  V120
    売上12 As Double                  '04/08/05  V120
    売上12不課税 As Double            '04/08/05  V120
    売上12課税 As Double              '04/08/05  V120
    売上12非課税 As Double            '04/08/05  V120
    売上13 As Double                  '04/08/05  V120
    売上13不課税 As Double            '04/08/05  V120
    売上13課税 As Double              '04/08/05  V120
    売上13非課税 As Double            '04/08/05  V120
    売上2 As Double
    売上21 As Double                  '04/08/05  V120
    売上21不課税 As Double            '04/08/05  V120
    売上21課税 As Double              '04/08/05  V120
    売上21非課税 As Double            '04/08/05  V120
    売上22 As Double                  '04/08/05  V120
    売上22不課税 As Double            '04/08/05  V120
    売上22課税 As Double              '04/08/05  V120
    売上22非課税 As Double            '04/08/05  V120
    売上23 As Double                  '04/08/05  V120
    売上23不課税 As Double            '04/08/05  V120
    売上23課税 As Double              '04/08/05  V120
    売上23非課税 As Double            '04/08/05  V120
    売上3 As Double
    売上31 As Double                  '04/08/05  V120
    売上31不課税 As Double            '04/08/05  V120
    売上31課税 As Double              '04/08/05  V120
    売上31非課税 As Double            '04/08/05  V120
    売上32 As Double                  '04/08/05  V120
    売上32不課税 As Double            '04/08/05  V120
    売上32課税 As Double              '04/08/05  V120
    売上32非課税 As Double            '04/08/05  V120
    売上33 As Double                  '04/08/05  V120
    売上33不課税 As Double            '04/08/05  V120
    売上33課税 As Double              '04/08/05  V120
    売上33非課税 As Double            '04/08/05  V120
    
    仕入額 As Double
    仕入1 As Double
    仕入11 As Double                  '05/04/04　V127
    仕入11不課税 As Double            '05/04/04  V127
    仕入11課税 As Double              '05/04/04  V127
    仕入11非課税 As Double            '05/04/04  V127
    仕入12 As Double                  '05/04/04  V127
    仕入12不課税 As Double            '05/04/04  V127
    仕入12課税 As Double              '05/04/04  V127
    仕入12非課税 As Double            '05/04/04  V127
    仕入13 As Double                  '05/04/04  V127
    仕入13不課税 As Double            '05/04/04  V127
    仕入13課税 As Double              '05/04/04  V127
    仕入13非課税 As Double            '05/04/04  V127
    仕入2 As Double
    仕入21 As Double                  '05/04/04  V127
    仕入21不課税 As Double            '05/04/04  V127
    仕入21課税 As Double              '05/04/04  V127
    仕入21非課税 As Double            '05/04/04  V127
    仕入22 As Double                  '05/04/04  V127
    仕入22不課税 As Double            '05/04/04  V127
    仕入22課税 As Double              '05/04/04  V127
    仕入22非課税 As Double            '05/04/04  V127
    仕入23 As Double                  '05/04/04  V127
    仕入23不課税 As Double            '05/04/04  V127
    仕入23課税 As Double              '05/04/04  V127
    仕入23非課税 As Double            '05/04/04  V127
    仕入3 As Double
    仕入31 As Double                  '05/04/04  V127
    仕入31不課税 As Double            '05/04/04  V127
    仕入31課税 As Double              '05/04/04  V127
    仕入31非課税 As Double            '05/04/04  V127
    仕入32 As Double                  '05/04/04  V127
    仕入32不課税 As Double            '05/04/04  V127
    仕入32課税 As Double              '05/04/04  V127
    仕入32非課税 As Double            '05/04/04  V127
    仕入33 As Double                  '05/04/04  V127
    仕入33不課税 As Double            '05/04/04  V127
    仕入33課税 As Double              '05/04/04  V127
    仕入33非課税 As Double            '05/04/04  V127
    
    回収実績 As Double                '04/08/10 V120
    M回収実績1 As Double              '06/02/01 V150
    M回収実績2 As Double              '06/02/01 V150
    投資回収実績 As Double            '05/07/25 V128
    M投資回収実績1 As Double          '06/02/01 V150
    M投資回収実績2 As Double          '06/02/01 V150
    未収回収実績 As Double            '06/02/01 V150
    M未収回収実績1 As Double          '06/02/01 V150
    M未収回収実績2 As Double          '06/02/01 V150
    非現金回収実績 As Double          '05/05/03 V127
    投資非現金回収実績 As Double      '06/02/01 V150
    未収非現金回収実績 As Double      '06/02/01 V150
    
    営業支払実績 As Double            '05/04/04 V127
    M営業支払実績1 As Double          '06/02/01 V150
    M営業支払実績2 As Double          '06/02/01 V150
    投資支払実績 As Double            '05/04/04 V127
    M投資支払実績1 As Double          '06/02/01 V150
    M投資支払実績2 As Double          '06/02/01 V150
    費用支払実績 As Double            '06/02/01 V150
    M費用支払実績1 As Double          '06/02/01 V150
    M費用支払実績2 As Double          '06/02/01 V150
    非現金支払実績 As Double          '05/05/03 V127
    投資非現金支払実績 As Double      '06/02/01 V150
    未払費用非現金支払実績 As Double  '06/02/01 V150
      
    粗利益 As Double
    粗利1 As Double
    粗利2 As Double
    粗利3 As Double
    
    '原価 As Double                     '05/04/04 V127
    '原価1 As Double                    '05/04/04 V127
    '原価2 As Double                    '05/04/04 V127
    '原価3 As Double                    '05/04/04 V127
    給与総額 As Double
    賞与額 As Double
    
    前受金等1 As Double               '06/02/01 V150
    前受金等2 As Double               '06/02/01 V150
    非現金前受金1 As Double           ' 07/04/17 V181
    非現金前受金2 As Double           ' 07/04/17 V181
    
    前払金等1 As Double               '06/02/01 V150
    前払金等2 As Double               '06/02/01 V150
    非現金前払金1 As Double           ' 07/04/17 V181
    非現金前払金2 As Double           ' 07/04/17 V181
    
    
    固定経費 As Double
    固定経費不課税 As Double
    固定経費課税 As Double
    固定経費非課税 As Double
    
    変動経費1 As Double
    変動経費1不課税 As Double
    変動経費1課税 As Double
    変動経費1非課税 As Double
    
    変動経費2 As Double
    変動経費2不課税 As Double
    変動経費2課税 As Double
    変動経費2非課税 As Double
    
    変動経費3 As Double
    変動経費3不課税 As Double
    変動経費3課税 As Double
    変動経費3非課税 As Double
    
    その他経費1 As Double
    その他経費1不課税 As Double
    その他経費1課税 As Double
    その他経費1非課税 As Double
    
    定期積金 As Double                '06/02/01 V150
    定期積金解約 As Double            '06/02/01 V150
    協力積立金 As Double              '06/02/01 V150
    協力積立金解約 As Double          '06/02/01 V150
    
    保険積立 As Double
    保険積立不課税 As Double
    保険積立課税 As Double
    保険積立非課税 As Double
    
    '保険積立解約 As Double         '06/03/11 V150
    '保険積立解約不課税 As Double   '06/03/11 V150
    '保険積立解約課税 As Double     '06/03/11 V150
    '保険積立解約非課税 As Double   '06/03/11 V150
    
    保険積立解約 As Double            ' 06/02/01 V150
    保険積立解約不課税 As Double
    保険積立解約課税 As Double
    保険積立解約非課税 As Double
    
    受取リベート As Double                ' 06/02/01 V150
    受取リベート不課税 As Double
    受取リベート課税 As Double
    受取リベート非課税 As Double
    
    支払リベート As Double                ' 06/02/01 V150
    支払リベート不課税 As Double
    支払リベート課税 As Double
    支払リベート非課税 As Double
    
    営業外収益 As Double
    営業外収益不課税 As Double
    営業外収益課税 As Double
    営業外収益非課税 As Double
    
    営業外費用 As Double
    営業外費用不課税 As Double
    営業外費用課税 As Double
    営業外費用非課税 As Double
    
    資産売却額 As Double              '05/07/23 V128
    資産売却額不課税 As Double        '05/07/23 V128
    資産売却額課税 As Double          '05/07/23 V128
    資産売却額非課税 As Double        '05/07/23 V128
    
    特別利益 As Double                '05/07/23 V128
    特別損失 As Double                '05/07/23 V128
    
    非現金利益 As Double              '05/11/21 V130
    非現金損失 As Double              '05/11/21 V130
    減価償却 As Double                  ' 07/04/20 V181
    
    非損益現金 As Double                ' 07/01/30 V180
    総債権調整額 As Double              ' 07/04/20 V181
    総債務調整額 As Double              ' 07/04/20 V181
    
    前受金 As Double                  '05/07/23 V128
    前払金 As Double                  '05/07/23 V128
    
    割引手形 As Double                '05/07/23 V128
    割引手形決済 As Double            '05/07/23 V128
    
    その他債権 As Double              ' 07/04/17 V181
    その他債権決済 As Double          ' 07/04/17 V181
    非現金その他債権 As Double        ' 07/04/17 V181
    非現金その他債権決済 As Double    ' 07/04/17 V181
    
    その他債務 As Double              ' 07/04/17 V181
    その他債務決済 As Double          ' 07/04/17 V181
    非現金その他債務 As Double        ' 07/04/17 V181
    非現金その他債務決済  As Double   ' 07/04/17 V181
    
'    前受金分売上額 As Double          '05/07/23 V128
'    前受金分売上額1 As Double         '05/07/23 V128
'    前受金分売上額1不課税 As Double   '05/07/23 V128
'    前受金分売上額1課税 As Double     '05/07/23 V128
'    前受金分売上額1非課税 As Double   '05/07/23 V128
'    前受金分売上額2 As Double         '05/07/23 V128
'    前受金分売上額2不課税 As Double   '05/07/23 V128
'    前受金分売上額2課税 As Double     '05/07/23 V128
'    前受金分売上額2非課税 As Double   '05/07/23 V128
'    前受金分売上額3 As Double         '05/07/23 V128
'    前受金分売上額3不課税 As Double   '05/07/23 V128
'    前受金分売上額3課税 As Double     '05/07/23 V128
'    前受金分売上額3非課税 As Double   '05/07/23 V128
'
'    前受金分粗利益 As Double          '05/07/23 V128
'    前受金分粗利益1 As Double         '05/07/23 V128
'    前受金分粗利益2 As Double         '05/07/23 V128
'    前受金分粗利益3 As Double         '05/07/23 V128
'
'    前渡金分仕入額 As Double          '05/07/23 V128
'    前渡金分仕入額1 As Double         '05/07/23 V128
'    前渡金分仕入額1不課税 As Double   '05/07/23 V128
'    前渡金分仕入額1課税 As Double     '05/07/23 V128
'    前渡金分仕入額1非課税 As Double   '05/07/23 V128
'    前渡金分仕入額2 As Double         '05/07/23 V128
'    前渡金分仕入額2不課税 As Double   '05/07/23 V128
'    前渡金分仕入額2課税 As Double     '05/07/23 V128
'    前渡金分仕入額2非課税 As Double   '05/07/23 V128
'    前渡金分仕入額3 As Double         '05/07/23 V128
'    前渡金分仕入額3不課税 As Double   '05/07/23 V128
'    前渡金分仕入額3課税 As Double     '05/07/23 V128
'    前渡金分仕入額3非課税 As Double   '05/07/23 V128
    
    販売データ有無 As Integer           '05/08/09 V129
    
    '減価償却 As Double
    支払利息 As Double
    
    取消フラグ As Integer
End Type

'------------------------< 基幹データ調整 >-------------------------
Type MAA910_基幹データ調整
    実績年月 As Date
    支店 As String                      '05/12/07 V130
    
    売上額 As Double
    
    売上1 As Double
    売上1不課税 As Double             '04/08/05  V120
    売上1課税 As Double               '04/08/05  V120
    売上1非課税 As Double             '04/08/05  V120
    回収サイト1 As Integer              '05/12/07 V130
    
    売上2 As Double
    売上2不課税 As Double            '04/08/05  V120
    売上2課税 As Double              '04/08/05  V120
    売上2非課税 As Double            '04/08/05  V120
    回収サイト2 As Integer             '05/12/07 V130
    
    売上3 As Double                  '04/08/05  V120
    売上3不課税 As Double            '04/08/05  V120
    売上3課税 As Double              '04/08/05  V120
    売上3非課税 As Double            '04/08/05  V120
    回収サイト3 As Integer             '05/12/07 V130
   
    仕入額 As Double
    
    仕入1 As Double                  '05/04/04　V127
    仕入1不課税 As Double            '05/04/04  V127
    仕入1課税 As Double              '05/04/04  V127
    仕入1非課税 As Double            '05/04/04  V127
    支払サイト1 As Integer             '05/12/07 V130
    
    仕入2 As Double                  '05/04/04  V127
    仕入2不課税 As Double            '05/04/04  V127
    仕入2課税 As Double              '05/04/04  V127
    仕入2非課税 As Double            '05/04/04  V127
    支払サイト2 As Integer             '05/12/07 V130
    
    仕入3 As Double                  '05/04/04  V127
    仕入3不課税 As Double            '05/04/04  V127
    仕入3課税 As Double              '05/04/04  V127
    仕入3非課税 As Double            '05/04/04  V127
    支払サイト3 As Integer             '05/12/07 V130
    
    粗利益 As Double
    粗利1 As Double
    粗利2 As Double
    粗利3 As Double
    
    回収実績 As Double                '04/08/10 V120
    M回収実績1 As Double              '06/02/01 V150
    M回収実績2 As Double              '06/02/01 V150
    投資回収実績 As Double            '05/07/25 V128
    M投資回収実績1 As Double          '06/02/01 V150
    M投資回収実績2 As Double          '06/02/01 V150
    未収回収実績 As Double            '06/02/01 V150
    M未収回収実績1 As Double          '06/02/01 V150
    M未収回収実績2 As Double          '06/02/01 V150
    非現金回収実績 As Double          '05/05/03 V127
    投資非現金回収実績 As Double      '06/02/01 V150
    未収非現金回収実績 As Double      '06/02/01 V150
    
    営業支払実績 As Double            '05/04/04 V127
    M営業支払実績1 As Double          '06/02/01 V150
    M営業支払実績2 As Double          '06/02/01 V150
    投資支払実績 As Double            '05/04/04 V127
    M投資支払実績1 As Double          '06/02/01 V150
    M投資支払実績2 As Double          '06/02/01 V150
    費用支払実績 As Double            '06/02/01 V150
    M費用支払実績1 As Double          '06/02/01 V150
    M費用支払実績2 As Double          '06/02/01 V150
    非現金支払実績 As Double          '05/05/03 V127
    投資非現金支払実績 As Double      '06/02/01 V150
    未払費用非現金支払実績 As Double  '06/02/01 V150
    
    給与総額 As Double
    賞与額 As Double
    
    前受金等1 As Double               '06/02/01 V150
    前受金等2 As Double               '06/02/01 V150
    非現金前受金1 As Double           ' 07/04/17 V181
    非現金前受金2 As Double           ' 07/04/17 V181
    前払金等1 As Double               '06/02/01 V150
    前払金等2 As Double               '06/02/01 V150
    非現金前払金1 As Double           ' 07/04/17 V181
    非現金前払金2 As Double           ' 07/04/17 V181
    
    固定経費 As Double
    固定経費不課税 As Double
    固定経費課税 As Double
    固定経費非課税 As Double
    
    変動経費1 As Double
    変動経費1不課税 As Double
    変動経費1課税 As Double
    変動経費1非課税 As Double
    
    変動経費2 As Double
    変動経費2不課税 As Double
    変動経費2課税 As Double
    変動経費2非課税 As Double
    
    変動経費3 As Double
    変動経費3不課税 As Double
    変動経費3課税 As Double
    変動経費3非課税 As Double
    
    その他経費1 As Double
    その他経費1不課税 As Double
    その他経費1課税 As Double
    その他経費1非課税 As Double
    
    定期積金 As Double                '06/02/01 V150
    定期積金解約 As Double            '06/02/01 V150
    協力積立金 As Double              '06/02/01 V150
    協力積立金解約 As Double          '06/02/01 V150
    
    保険積立 As Double
    保険積立不課税 As Double
    保険積立課税 As Double
    保険積立非課税 As Double
    
    保険積立解約 As Double            ' 06/02/01 V150
    保険積立解約不課税 As Double
    保険積立解約課税 As Double
    保険積立解約非課税 As Double
    
    受取リベート As Double                ' 06/02/01 V150
    受取リベート不課税 As Double
    受取リベート課税 As Double
    受取リベート非課税 As Double
    
    支払リベート As Double                ' 06/02/01 V150
    支払リベート不課税 As Double
    支払リベート課税 As Double
    支払リベート非課税 As Double
    
    営業外収益 As Double
    営業外収益不課税 As Double
    営業外収益課税 As Double
    営業外収益非課税 As Double
    
    営業外費用 As Double
    営業外費用不課税 As Double
    営業外費用課税 As Double
    営業外費用非課税 As Double
    
    資産売却額 As Double              '05/07/23 V128
    資産売却額不課税 As Double        '05/07/23 V128
    資産売却額課税 As Double          '05/07/23 V128
    資産売却額非課税 As Double        '05/07/23 V128
    
    特別利益 As Double                '05/07/23 V128
    特別損失 As Double                '05/07/23 V128
    
    非現金利益 As Double              '05/11/21 V130
    非現金損失 As Double              '05/11/21 V130
    減価償却 As Double                  ' 07/04/20 V181
    
    非損益現金 As Double                ' 07/01/30 V180
    総債権調整額 As Double              ' 07/04/20 V181
    総債務調整額 As Double              ' 07/04/20 V181
    
    割引手形 As Double                '05/12/10 V130
    割引手形決済 As Double            '05/12/10 V130
        
    その他債権 As Double              ' 07/04/17 V181
    その他債権決済 As Double          ' 07/04/17 V181
    非現金その他債権 As Double        ' 07/04/17 V181
    非現金その他債権決済 As Double    ' 07/04/17 V181
    
    その他債務 As Double              ' 07/04/17 V181
    その他債務決済 As Double          ' 07/04/17 V181
    非現金その他債務 As Double        ' 07/04/17 V181
    非現金その他債務決済  As Double   ' 07/04/17 V181
    
    取消フラグ As Integer
End Type
'

'------------------------< 本部経費振替 >-------------------------　　5/11/16 V130
Type MAA910_本部経費振替
    実績年月 As Date
    支店 As String
    
    回収実績 As Double                '04/08/10 V120
    M回収実績1 As Double              '06/02/01 V150
    M回収実績2 As Double              '06/02/01 V150
    投資回収実績 As Double            '05/07/25 V128
    M投資回収実績1 As Double          '06/02/01 V150
    M投資回収実績2 As Double          '06/02/01 V150
    未収回収実績 As Double            '06/02/01 V150
    M未収回収実績1 As Double          '06/02/01 V150
    M未収回収実績2 As Double          '06/02/01 V150
    非現金回収実績 As Double          '05/05/03 V127
    投資非現金回収実績 As Double      '06/02/01 V150
    未収非現金回収実績 As Double      '06/02/01 V150
    
    営業支払実績 As Double            '05/04/04 V127
    M営業支払実績1 As Double          '06/02/01 V150
    M営業支払実績2 As Double          '06/02/01 V150
    投資支払実績 As Double            '05/04/04 V127
    M投資支払実績1 As Double          '06/02/01 V150
    M投資支払実績2 As Double          '06/02/01 V150
    費用支払実績 As Double            '06/02/01 V150
    M費用支払実績1 As Double          '06/02/01 V150
    M費用支払実績2 As Double          '06/02/01 V150
    非現金支払実績 As Double          '05/05/03 V127
    投資非現金支払実績 As Double      '06/02/01 V150
    未払費用非現金支払実績 As Double  '06/02/01 V150
    
    給与総額 As Double
    賞与額 As Double
    
    前受金等1 As Double               '06/02/01 V150
    前受金等2 As Double               '06/02/01 V150
    非現金前受金1 As Double           ' 07/04/17 V181
    非現金前受金2 As Double           ' 07/04/17 V181
    前払金等1 As Double               '06/02/01 V150
    前払金等2 As Double               '06/02/01 V150
    非現金前払金1 As Double           ' 07/04/17 V181
    非現金前払金2 As Double           ' 07/04/17 V181
    
    固定経費 As Double
    固定経費不課税 As Double
    固定経費課税 As Double
    固定経費非課税 As Double
    
    変動経費1 As Double
    変動経費1不課税 As Double
    変動経費1課税 As Double
    変動経費1非課税 As Double
    
    変動経費2 As Double
    変動経費2不課税 As Double
    変動経費2課税 As Double
    変動経費2非課税 As Double
    
    変動経費3 As Double
    変動経費3不課税 As Double
    変動経費3課税 As Double
    変動経費3非課税 As Double
    
    その他経費1 As Double
    その他経費1不課税 As Double
    その他経費1課税 As Double
    その他経費1非課税 As Double
    
    定期積金 As Double                '06/02/01 V150
    定期積金解約 As Double            '06/02/01 V150
    協力積立金 As Double              '06/02/01 V150
    協力積立金解約 As Double          '06/02/01 V150
    
    保険積立 As Double
    保険積立不課税 As Double
    保険積立課税 As Double
    保険積立非課税 As Double
    
    保険積立解約 As Double            ' 06/02/01 V150
    保険積立解約不課税 As Double
    保険積立解約課税 As Double
    保険積立解約非課税 As Double
    
    受取リベート As Double                ' 06/02/01 V150
    受取リベート不課税 As Double
    受取リベート課税 As Double
    受取リベート非課税 As Double
    
    支払リベート As Double                ' 06/02/01 V150
    支払リベート不課税 As Double
    支払リベート課税 As Double
    支払リベート非課税 As Double
    
    営業外収益 As Double
    営業外収益不課税 As Double
    営業外収益課税 As Double
    営業外収益非課税 As Double
    
    営業外費用 As Double
    営業外費用不課税 As Double
    営業外費用課税 As Double
    営業外費用非課税 As Double
    
    資産売却額 As Double              '05/07/23 V128
    資産売却額不課税 As Double        '05/07/23 V128
    資産売却額課税 As Double          '05/07/23 V128
    資産売却額非課税 As Double        '05/07/23 V128
    
    特別利益 As Double                '05/07/23 V128
    特別損失 As Double                '05/07/23 V128
    
    非現金利益 As Double              '05/11/21 V130
    非現金損失 As Double              '05/11/21 V130
    減価償却 As Double                  ' 07/04/20 V181
    
    非損益現金 As Double                ' 07/01/30 V180
    総債権調整額 As Double              ' 07/04/20 V181
    総債務調整額 As Double              ' 07/04/20 V181
    
    前受金 As Double                  '05/07/23 V128
    前払金 As Double                  '05/07/23 V128
    
    割引手形 As Double                '05/07/23 V128
    割引手形決済 As Double            '05/07/23 V128
        
    その他債権 As Double              ' 07/04/17 V181
    その他債権決済 As Double          ' 07/04/17 V181
    非現金その他債権 As Double        ' 07/04/17 V181
    非現金その他債権決済 As Double    ' 07/04/17 V181
    
    その他債務 As Double              ' 07/04/17 V181
    その他債務決済 As Double          ' 07/04/17 V181
    非現金その他債務 As Double        ' 07/04/17 V181
    非現金その他債務決済  As Double   ' 07/04/17 V181
    
    取消フラグ As Integer
End Type
'

'------------------------< 売上実績販売 >-------------------------  5/8/9 V129
Type MAA910_売上実績販売
    データ種別 As String                '05/08/09 V129
    対象年月 As Date                    '05/08/09 V129
    計上年月 As Date                    '05/08/09 V129
    支店 As String                      '05/08/09 V129
    計上区分 As String                  '05/08/09 V129
    売上分類 As Integer                 '05/08/09 V129
    不課税金額 As Double              '05/08/09 V129
    課税金額 As Double                '05/08/09 V129
    非課税金額 As Double              '05/08/09 V129
    粗利金額 As Double                '05/08/09 V129
    サイト As Integer                   '05/08/09 V129
    取消フラグ As Integer               '05/08/09 V129
End Type                                '05/08/09 V129
'

'------------------------< 受注発注 >-------------------------  5/8/9 V129
'Type MAA910_受注発注
'    データ種別 As String                '05/08/09 V129
'    対象年月 As Date                    '05/08/09 V129
'    計上年月 As Date                    '05/08/09 V129
'    支店 As String                      '05/08/09 V129
'    計上区分 As String                  '05/08/09 V129
'    売上分類 As Integer                 '05/08/09 V129
'    不課税金額 As Double              '05/08/09 V129
'    課税金額 As Double                '05/08/09 V129
'    非課税金額 As Double              '05/08/09 V129
'    粗利金額 As Double                '05/08/09 V129
'    サイト As Integer                   '05/08/09 V129
'    取消フラグ As Integer               '05/08/09 V129
'End Type                                '05/08/09 V129
'

'------------------------< 設備計画 >-------------------------
Type MAA910_設備計画
    設備番号 As String
    設備計画番号 As String
    '設備リストラ番号 As String
    SM区分 As Integer
    残高移行区分 As Long
    設備名 As String
    設備購入年月日 As Variant           ' 07/01/30 V180
    設備年月 As Variant
    減価償却除外開始年月(20) As Variant ' 08/07/09 V188
    減価償却除外終了年月(20) As Variant ' 08/07/09 V188
    償却最終年月 As Variant
    償却区分 As String
    減価償却費区分 As String
    資産区分 As String
    設備金額 As Double
    支払サイト As Integer               ' 07/01/30 V180
    償却年数 As Integer
    残存率 As Double                    ' 07/01/30 V180
    調整償却額 As Double                ' 07/02/05 V180
    調整償却額2 As Double               ' 08/08/23 V188
    特別償却1年次額 As Double           ' 07/01/30 V180
    特別償却2年次額 As Double           ' 07/01/30 V180
    特別償却3年次額 As Double           ' 07/01/30 V180
    設備リストラ番号 As String          ' 07/01/30 V180
    資産売却年月日 As Variant           ' 07/01/30 V180
    資産売却額 As Double                ' 07/01/30 V180
    回収サイト As Integer               ' 07/01/30 V180
    売上課税区分 As String              ' 07/02/06 V180
    手入力フラグ As Integer             ' 07/01/30 V180
    修正不可F As Integer                ' 07/02/18 V180
    課税区分 As String
    '廃棄年月 As Variant
    'リストラ廃棄年月 As Variant
    取消フラグ As Integer
End Type

Type MAA910_設備計画テーブル
    設備番号 As String
    償却年月 As Variant
    減価償却控除開始年月 As Variant     ' 08/07/15 V188
    減価償却控除終了年月 As Variant     ' 08/07/15 V188
    期首簿価 As Double
    償却金額 As Double
    残存金額 As Double
    調整償却額  As Double               ' 07/02/05 V180
    調整償却額2 As Double               ' 08/09/01 V188
    特別償却額 As Double                ' 07/01/30 V180
    売却額 As Double                    ' 07/01/30 V180
    売却益 As Double                    ' 07/01/30 V180
    売却損 As Double                    ' 07/01/30 V180
End Type
'

'---------------------------------------------------------------------------------
